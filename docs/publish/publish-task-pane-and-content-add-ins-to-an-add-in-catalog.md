# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>Publication de compléments du volet Office et de contenu dans un catalogue SharePoint

Un catalogue de compléments est une collection de sites dédiée dans une application web SharePoint ou une location SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office et SharePoint. Pour rendre les compléments Office accessibles aux utilisateurs, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue de compléments pour leur organisation. Lorsqu’un administrateur enregistre un catalogue de compléments en tant que catalogue approuvé, les utilisateurs peuvent insérer le complément à partir de l’interface utilisateur d’insertion dans une application cliente Office.

**Remarques importantes :** 

- les catalogues de compléments sur SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud `VersionOverrides` du [manifeste de complément](../overview/add-in-manifests.md), comme les commandes de complément.

- Si vous ciblez un environnement de cloud ou hybride, nous vous recommandons d’[utiliser un déploiement centralisé via le centre d’administration Office 365](publish/centralized-deployment.md) pour publier vos compléments.

- Les catalogues SharePoint ne sont pas pris en charge dans Office 2016 pour Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à l’[Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx).   

## <a name="set-up-an-add-in-catalog"></a>Configuration d’un catalogue de compléments

Suivez les étapes décrites dans l’une des sections suivantes pour configurer un catalogue de compléments sur SharePoint ou Office 365.

### <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a>Configuration d’un catalogue de compléments sur SharePoint

1. Accédez au **site Administration centrale** (**Démarrer** > **Tous les programmes** > **Produits Microsoft SharePoint 2013** > **Administration centrale SharePoint 2013**).
    
2. Dans le volet Office de gauche, cliquez sur  **Compléments**.
    
3. Sur la page  **Compléments**, sous  **Gestion des compléments**, choisissez  **Gérer le catalogue de compléments**.
    
4. Sur la page  **Gérer le catalogue de compléments**, vérifiez que vous avez sélectionné l’application web appropriée dans  **Sélecteur d’applications web**.
    
5. Choisissez  **Afficher les paramètres du site**.
    
6. Sur la page  **Paramètre du site**, choisissez  **Administrateurs de collections de sites** pour spécifier les administrateurs de collection de sites, puis choisissez **OK**.
    
7. Pour accorder des autorisations de site aux utilisateurs, choisissez  **Autorisations de site**, puis choisissez  **Accorder des autorisations**.
    
8. Dans la boîte de dialogue  **Partager le site de catalogue d’applications**, spécifiez des utilisateurs de site, définissez les autorisations appropriées pour ces derniers, puis éventuellement d’autres options, puis choisissez  **Partager**.
    
9. Pour ajouter un complément au catalogue de compléments Office, choisissez **Compléments Office**.

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a>Configuration d’un catalogue de compléments sur Office 365

1. Sur la page Centre d’administration Office 365, sélectionnez **Administrateur**, puis **SharePoint**.
    
2. Dans le volet Office situé à gauche, cliquez sur  **Compléments**.
    
3. Sur la page  **Compléments**, cliquez sur  **Catalogue de compléments**.
    
4. Sur la page  **Site de catalogue de compléments**, cliquez sur  **OK** pour accepter l’option par défaut et créer un site de catalogue de compléments.
    
5. Sur la page  **Créer une collection de sites de catalogue de compléments**, indiquez le titre de votre site de catalogue de compléments.
    
6. Spécifiez l’adresse du site web.
    
7. Définissez le **quota de stockage** sur la valeur la plus petite possible (110 actuellement). Vous allez installer uniquement des packages de compléments sur cette collection de sites et ils sont très petits.
    
8. Définissez le **quota de ressources du serveur** sur 0 (zéro). (Le quota de ressources du serveur est lié à la limitation des solutions bac à sable (sandbox) dont les performances sont médiocres, mais vous n’installerez aucune solution bac à sable (sandbox) sur votre site de catalogue de compléments.)
    
9. Sélectionnez **OK**.
    
10. Pour ajouter un complément au site du catalogue de compléments, accédez au site que vous venez de créer. Dans le volet de navigation gauche, sélectionnez **Compléments Office**, puis, pour charger un fichier manifeste de complément Office, sélectionnez **Nouveau complément**.

## <a name="publish-an-add-in-to-an-add-in-catalog"></a>Publication d’un complément dans un catalogue de compléments

Pour publier un complément dans un catalogue de compléments, procédez comme suit.

1. Accédez au catalogue de compléments :

    1- Ouvrez la page principale de l’Administration centrale de SharePoint.
    
    2- Sélectionnez **Compléments**.
    
    3- Sélectionnez **Gérer le catalogue de compléments**.
    
    4- Sélectionnez le lien fourni, puis choisissez **Compléments Office** dans la barre de navigation située à gauche.
    
2. Sélectionnez le lien **Cliquer pour ajouter un nouvel élément**.
    
3. Choisissez **Parcourir**, puis spécifiez le [manifeste](../overview/add-in-manifests.md) à télécharger.
    
    Les compléments de contenu et de volet Office de ce catalogue sont désormais disponibles dans la boîte de dialogue **Compléments Office**. Pour y accéder, choisissez **Mes compléments** sous l’onglet **Insérer**, puis choisissez **MON ORGANISATION**.

## <a name="end-user-experience-with-the-add-in-catalog"></a>Expérience des utilisateurs finaux avec le catalogue des compléments

Les utilisateurs finaux peuvent accéder au catalogue des compléments dans une application Office en procédant comme suit :

1. Dans l’application Office, accédez à **Fichier**  >  **Options**  >  **Centre de gestion de la confidentialité**  >  **Paramètres du centre de gestion de la confidentialité**  >  **Catalogues de compléments approuvés**.
    
2. Spécifiez l’URL de la _collection de sites SharePoint parente_ du catalogue de compléments. Par exemple, si l’URL du catalogue de compléments Office est :
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    Spécifiez simplement l’URL de la collection de sites parente :
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. Fermez puis rouvrez l’application Office. Le catalogue de compléments est disponible dans la boîte de dialogue **Compléments Office**.

Par ailleurs, un administrateur peut spécifier un catalogue de compléments Office sur SharePoint à l’aide d’une stratégie de groupe. Pour plus d’informations, reportez-vous à la section relative à l’[utilisation d’une stratégie de groupe pour gérer la manière dont les utilisateurs peuvent installer et utiliser des compléments Office](https://technet.microsoft.com/en-us/library/jj219429.aspx#BKMK_GP) sur TechNet.

