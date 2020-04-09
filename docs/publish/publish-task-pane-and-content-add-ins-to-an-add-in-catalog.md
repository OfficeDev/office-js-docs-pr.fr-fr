---
title: Publier des compléments du volet Office et de contenu dans un catalogue d’applications SharePoint
description: Pour rendre les compléments Office accessibles aux utilisateurs, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue d’applications pour leur organisation.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5557dd31e829fac2c2dbd421200da46a5c3b9b99
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185588"
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

Pour créer le catalogue d’applications SharePoint, suivez les instructions de la section [Configurer le site de catalogue d'applications pour une application Web](/sharepoint/administration/manage-the-app-catalog).

Une fois que vous avez créé le catalogue d’applications, suivez les étapes pour [publier un complément Office](#publish-an-office-add-in).

### <a name="to-create-an-app-catalog-on-office-365"></a>Pour créer catalogue d’applications Office 365

Pour créer le catalogue d’applications SharePoint, suivez les instructions de [la rubrique créer la collection de sites de catalogue d’applications](/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection). Une fois que vous avez créé le catalogue d’applications, suivez les étapes de la section suivante pour publier un complément Office.

## <a name="publish-an-office-add-in"></a>Publier un complément Office

Suivez les étapes décrites dans l’une des sections suivantes pour publier un complément Office dans un catalogue d’applications avec Office 365 ou avec SharePoint Server local.

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a>Pour publier un complément Office dans un catalogue d’applications SharePoint sur Office 365

1. Accédez à la [Page de sites actifs du nouveau Centre d’administration SharePoint](https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true) et connectez-vous à l’aide d’un compte disposant des [autorisations d’administrateur](/sharepoint/sharepoint-admin-role) pour votre organisation.

>[!NOTE]
>Si vous avez Office 365 Allemagne, [connectez-vous au Centre d’administration Microsoft 365](https://go.microsoft.com/fwlink/p/?linkid=848041), puis recherchez dans le Centre d’administration SharePoint et ouvrez la page Autres fonctionnalités. <br>Si vous avez Office 365 géré par 21Vianet (en Chine), [connectez-vous au Centre d’administration Microsoft 365](https://go.microsoft.com/fwlink/p/?linkid=850627), puis recherchez dans le Centre d’administration SharePoint et ouvrez la page Autres fonctionnalités.
 
2. Ouvrez le site de catalogue d’applications en sélectionnant son URL dans la colonne URL. 

>[!NOTE]
>Si vous venez de créer le site de catalogue d’applications dans la section précédente, la configuration du site peut prendre quelques minutes.

3. Choisissez **Distribuer des applications pour Office**.
4. Dans la page **Applications pour Office**, cliquez sur **Nouveau**.
5. Dans la boîte de dialogue **Ajouter un document**, sélectionnez le bouton **Choisir un fichier**.
6. Recherchez et spécifiez le fichier [manifeste](../develop/add-in-manifests.md) à télécharger, puis sélectionnez **Ouvrir**.
7. Dans la boîte de dialogue **Ajouter un document**, cliquez sur **OK**.

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
5. Choisissez un complément Office, puis sélectionnez **Ajouter**.

Par ailleurs, un administrateur peut spécifier un catalogue d’applications sur SharePoint à l’aide d’une stratégie de groupe. Les paramètres de stratégie pertinents sont disponibles sur la page [Modèles de fichiers administratifs (ADMX/ADML) pour Office 365 ProPlus, Office 2019 et Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) et se trouvent sous **User Configuration\Policies\Administrative Templates\Microsoft Office 2016 \ sécurité Settings\Trust Center\Trusted Catalogs**.
