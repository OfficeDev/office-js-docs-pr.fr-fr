# <a name="deploy-and-publish-your-office-add-in"></a>Déploiement et publication de votre complément Office

Vous pouvez utiliser l’une des méthodes pour déployer votre complément Office à des fins de test ou de distribution auprès des utilisateurs.

|**Méthode**|**Use...**|
|:---------|:------------|
|[Chargement de version test](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Dans le cadre de votre processus de développement, pour tester l’exécution de votre complément sur Windows, Office Online, iPad ou Mac.|
|[Déploiement centralisé](centralized-deployment.md)|Dans un environnement de cloud ou hybride, utilisez cette méthode pour distribuer votre complément auprès des utilisateurs de votre organisation à l’aide du Centre d’administration Office 365.|
|[Catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Dans un environnement local, pour distribuer votre complément auprès des utilisateurs de votre organisation.|
|[Office Store](https://dev.office.com/officestore/docs/submit-to-the-office-store)|Pour distribuer publiquement votre complément auprès des utilisateurs.|
|[Serveur Exchange](#outlook-add-in-deployment)|Dans un environnement local ou en ligne, pour distribuer des compléments Outlook à des utilisateurs.|

> [!NOTE]
>  si vous envisagez d’envoyer votre complément à l’Office Store, assurez-vous que vous respectez les [stratégies de validation de l’Office Store](https://msdn.microsoft.com/fr-fr/library/jj220035.aspx). Par exemple, pour obtenir la validation, votre complément doit fonctionner sur toutes les plateformes qui prennent en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability)).

## <a name="deployment-options-by-office-host"></a>Options de déploiement par l’hôte Office

Les options de déploiement disponibles dépendent de l’hôte Office que vous ciblez et du type de complément que vous créez.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Options de déploiement pour les compléments Word, Excel et PowerPoint

| Point d’extension | Chargement de version test | Centre d’administration Office 365 |Office Store| Catalogue SharePoint*  |
|:----------------|:------------|:-------------------|:--------------------------------|:-------------|
| Contenu         | X           | X                  | X                               | X|
| Volet Office       | X           | X                  | X                               | X|
| Commande         | X           | X                  | X                               |  |

&#42; Les catalogues SharePoint ne prennent pas en charge Office 2016 pour Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Options de déploiement pour les compléments Outlook

| Point d’extension | Chargement de version test | Serveur Exchange | Office Store |
|:---------|:------------|:----------------|:-------------|
| Application de messagerie | X           | X               | X            |
| Commande  | X           | X               | X            |

## <a name="deployment-methods"></a>Méthodes de déploiement

Les sections suivantes fournissent des informations supplémentaires sur les méthodes de déploiement les plus fréquemment utilisées pour distribuer des compléments Office aux utilisateurs au sein d’une organisation.

Pour plus d’informations sur l’acquisition, l’insertion et l’exécution des compléments par les utilisateurs finaux, consultez l’article relatif aux [premiers pas de l’utilisation de votre complément Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

### <a name="centralized-deployment-via-the-office-365-admin-center"></a>Déploiement centralisé via le centre d’administration Office 365 

Le centre d’administration Office 365 permet aux administrateurs de déployer facilement des compléments Office auprès d’utilisateurs et de groupes dans leur organisation. Les compléments déployés via le centre d’administration sont disponibles pour les utilisateurs directement dans leurs applications Office, sans qu’aucune configuration client ne soit requise. Vous pouvez utiliser le déploiement centralisé pour déployer des compléments internes, ainsi que des compléments fournis par des éditeurs de logiciels indépendants.

Pour plus d’informations, reportez-vous à [Publication des compléments Office à l’aide du déploiement centralisé via le centre d’administration Office 365](centralized-deployment.md).

### <a name="sharepoint-catalog-deployment"></a>Déploiement d’un catalogue SharePoint

Un catalogue de compléments SharePoint est une collection de sites spéciale que vous pouvez créer pour héberger des compléments Word, Excel et PowerPoint. Les catalogues SharePoint ne prennent pas en charge les nouvelles fonctionnalités de complément mises en œuvre dans le nœud `VersionOverrides` du manifeste, y compris les commandes de complément. Nous vous recommandons d’utiliser un déploiement centralisé via le centre d’administration, si possible. Par défaut, les commandes de complément déployées via un catalogue SharePoint s’ouvrent dans un volet des tâches.

Si vous déployez des compléments dans un environnement local, utilisez un catalogue SharePoint. Pour obtenir des détails, voir l’article sur la [publication de compléments du volet des tâches et de contenu dans un catalogue SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Remarque :** les catalogues SharePoint ne prennent pas en charge Office 2016 pour Mac. Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à l’[Office Store]. 

### <a name="outlook-add-in-deployment"></a>Déploiement de compléments Outlook

Pour des environnements en ligne et locaux qui n’utilisent pas le service d’identité Azure AD, vous pouvez déployer des compléments Outlook via le serveur Exchange. 

Le déploiement de compléments Outlook nécessite :

- Office 365, Exchange Online ou Exchange Server 2013 ou version ultérieure
- Outlook 2013 ou une version ultérieure

Pour affecter des compléments à des clients, utilisez le centre d’administration Exchange pour télécharger un manifeste directement, à partir d’un fichier ou d’une URL, ou ajoutez un complément à partir de l’Office Store. Pour affecter des compléments à des utilisateurs individuels, vous devez utiliser Exchange PowerShell. Pour plus d’informations, voir [Installation ou suppression de compléments Outlook pour votre organisation](https://technet.microsoft.com/fr-fr/library/jj943752(v=exchg.150).aspx) sur TechNet.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md) 
- [Envoyer à l’Office Store][Office Store]
- [Instructions de conception pour les compléments Office](../design/add-in-design)
- 
  [Création de compléments efficaces pour l’Office Store](https://msdn.microsoft.com/fr-fr/library/jj635874.aspx)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)

[Office Store]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
