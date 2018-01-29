
# <a name="host-an-office-add-in-on-microsoft-azure"></a>Héberger un complément pour Office sur Microsoft Azure

Le complément Office le plus simple est constitué d’un fichier manifeste XML et d’une page HTML. Le fichier manifeste XML décrit les caractéristiques du complément, telles que son nom, les applications clientes Office dans lesquelles il peut s’exécuter et l’URL de la page HTML du complément. La page HTML est contenue dans une application web avec laquelle les utilisateurs interagissent lorsqu’ils installent et exécutent votre complément au sein d’une application cliente Office. Vous pouvez héberger l’application web d’un complément Office sur n’importe quelle plateforme d’hébergement web, y compris Azure.

Cet article décrit comment déployer une application web de complément sur Azure et [charger une version test du complément](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) pour le tester dans une application cliente Office.

## <a name="prerequisites"></a>Conditions préalables 

1. Installez [Visual Studio 2017](https://www.visualstudio.com/downloads) et choisissez d’inclure la charge de travail de **développement Azure**.

    >**Remarque :** Si vous avez déjà installé Visual Studio 2017, [utilisez le programme d’installation Visual Studio](https://docs.microsoft.com/fr-fr/visualstudio/install/modify-visual-studio) pour vous assurer que la charge de travail de **développement Azure** est installée. 

2. Installez Office 2016. 
    
     >**Remarque :** Si vous n’avez pas encore Office 2016, vous pouvez vous [inscrire pour une version d’évaluation gratuite d’un mois](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).

3.  Obtenez un abonnement Azure.
    
     >**Remarque :** Si vous n’avez pas encore d’abonnement Azure, vous pouvez [en obtenir un dans le cadre de votre abonnement MSDN](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/) ou vous [inscrire pour obtenir une version d’évaluation gratuite](https://azure.microsoft.com/en-us/pricing/free-trial). 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a>Étape 1 : Créer un dossier partagé pour héberger le fichier manifeste XML de votre complément

1. Ouvrez l’explorateur de fichiers sur votre ordinateur de développement.
    
2. Cliquez avec le bouton droit de la souris sur le lecteur C:\, puis choisissez **Nouveau** > **Dossier**.
    
3. Nommez le nouveau dossier AddinManifests.
    
4. Cliquez avec le bouton droit de la souris sur le dossier AddinManifests, puis choisissez **Partager avec** > **Des personnes spécifiques**.
    
5. Dans **Partage de fichiers**, sélectionnez la flèche déroulante vers le bas, puis choisissez **Tout le monde** > **Ajouter** > **Partager**.
    
> **Remarque :** Dans cette procédure, vous utilisez un partage de fichiers local en tant que catalogue approuvé où vous allez stocker le fichier manifeste XML du complément. Dans un scénario réel, vous pouvez choisir à la place de [déployer le fichier manifeste XML dans un catalogue SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou de [publier le complément dans l’Office Store](https://dev.office.com/officestore/docs/submit-to-the-office-store).

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a>Étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés

1.  Démarrez Word 2016 et créez un document.

    >**Remarque :** Bien que cet exemple utilise Word 2016, vous pouvez utiliser n’importe quelle application Office qui prend en charge des compléments Office comme Excel, Outlook, PowerPoint ou Project 2016.
    
2.  Choisissez **Fichier**  >  **Options**.
    
3.  Dans la boîte de dialogue **Options Word**, choisissez **Centre de gestion de la confidentialité**, puis **Paramètres du Centre de gestion de la confidentialité**. 
    
4.  Dans la boîte de dialogue **Centre de gestion de la confidentialité**, choisissez **Catalogues de compléments approuvés**. Saisissez le chemin d’accès UNC (Universal Naming Convention) pour le partage de fichiers que vous avez créé précédemment en tant qu’**URL du catalogue** (par exemple, \\\NomDeVotreOrdinateur\AddinManifests), puis choisissez **Ajouter un catalogue**. 
    
5. Activez la case **Afficher dans le menu**. 

    >**Remarque :** Lorsque vous stockez un fichier manifeste XML de complément sur un partage qui est défini comme un catalogue de compléments web approuvés, le complément apparaît sous **Dossier partagé** dans la boîte de dialogue **Compléments Office** lorsque l’utilisateur accède à l’onglet **Insérer** dans le ruban et choisit **Mes compléments**.

6. Fermez Word 2016.

## <a name="step-3-create-a-web-app-in-azure"></a>Étape 3 : Créer une application web dans Azure

Créez une application web vide dans Azure en utilisant [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) ou le [portail Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).

### <a name="using-visual-studio-2017"></a>Utilisation de Visual Studio 2017

Pour créer l’application web à l’aide de Visual Studio 2017, procédez comme suit.

1. Dans Visual Studio, dans le menu **Affichage**, sélectionnez **Explorateur de serveurs**. Cliquez avec le bouton droit de la souris sur **Azure** et choisissez **Se connecter à un abonnement Microsoft Azure**. Suivez les instructions pour vous connecter à votre abonnement Azure.
    
2. Dans Visual Studio, dans **Explorateur de serveurs**, développez **Azure**, cliquez avec le bouton droit de la souris sur **App Service**, puis choisissez **Créer un App Service**.
    
3. Dans la boîte de dialogue **Créer App Service**, indiquez les informations suivantes :
    
      - Entrez un **nom d’application web** unique pour votre site. Azure vérifie que le nom du site est unique dans le domaine azurewebsites.net.

      - Choisissez l’**abonnement** à utiliser pour créer ce site.

      - Choisissez le **groupe de ressources** pour votre site. Si vous créez un groupe, vous devez également le nommer.
    
      - Choisissez le **plan de service d'applications** à utiliser pour créer ce site. Si vous créez un plan, vous devez également le nommer.
       
      - Sélectionnez **Créer**.

    La nouvelle application web s’affiche dans **Explorateur de serveurs** sous **Azure** >> **App Service** >> (le groupe de ressources choisi).
    
4. Cliquez avec le bouton droit de la souris sur la nouvelle application web, puis choisissez **Afficher dans le navigateur**. Votre navigateur s’ouvre et affiche une page web avec le message « Votre service d’application a été créé. ».
    
5. Dans la barre d’adresse du navigateur, modifiez l’URL de l’application web pour qu’elle utilise le protocole HTTPS et appuyez sur **Entrée** pour confirmer que le protocole HTTPS est activé. Le modèle de complément Office nécessite des compléments pour utiliser le protocole HTTPS.
    
### <a name="using-the-azure-portal"></a>Utilisation du portail Azure

Pour créer l’application web à l’aide du portail Azure, procédez comme suit.

1. Connectez-vous au [portail Azure](https://portal.azure.com/) à l’aide de vos informations d’identification Azure.
    
2. Choisissez **Nouveau** > **Web + mobile** > **Application web**. 

3. Dans la boîte de dialogue **Créer une application web**, renseignez ces informations :
    
      - Entrez un **nom d’application** unique pour votre site. Azure vérifie que le nom du site est unique dans le domaine apps.net azureweb.

      - Choisissez l’**abonnement** à utiliser pour créer ce site.

      - Choisissez le **groupe de ressources** pour votre site. Si vous créez un groupe, vous devez également le nommer.

      - Choisissez le **système d’exploitation** de votre site.
    
      - Choisissez le **plan de service d’applications** à utiliser pour créer ce site. Si vous créez un plan, vous devez également le nommer.
       
      - Sélectionnez **Créer**.

4. Choisissez **Notifications** (l’icône représentant une cloche qui se trouve sur le bord supérieur du portail Azure), puis choisissez la notification **Déploiements réussis** pour ouvrir la page **Vue d’ensemble** du site dans le portail Azure.

    >**Remarque :** La notification passera de **Déploiement en cours** à **Déploiements réussis** quand le déploiement du site sera terminé.

5. Dans la section **Essentials** de la page **Vue d’ensemble** du site dans le portail Azure, sélectionnez l’URL qui s’affiche sous **URL**. Votre navigateur s’ouvre et affiche une page web avec le message « Votre service d’application a été créé. ». 
    
6. Dans la barre d’adresse du navigateur, modifiez l’URL de l’application web pour qu’elle utilise le protocole HTTPS et appuyez sur **Entrée** pour confirmer que le protocole HTTPS est activé. Le modèle de complément Office nécessite des compléments pour utiliser le protocole HTTPS.    

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a>Étape 4 : Créer un complément Office dans Visual Studio

1. Démarrez Visual Studio en tant qu’administrateur.
    
2. Choisissez **Fichier** > **Nouveau** > **Projet**.
    
3. Sous **Modèles**, développez **Visual C#** (ou **Visual Basic**), développez **Office/SharePoint** et choisissez **Compléments**.
    
4. Choisissez **Complément Word web**, puis cliquez sur **OK** pour accepter les paramètres par défaut.
       
Visual Studio crée un complément Word de base que vous pourrez publier tel quel, sans apporter de modifications à son projet web.

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a>Étape 5 : Publier votre application web de complément Office sur Azure

1. Avec votre projet de complément ouvert dans Visual Studio, développez le nœud de solution dans l’**explorateur de solutions** pour voir les deux projets pour la solution.
    
2. Cliquez avec le bouton droit de la souris sur le projet web, puis choisissez **Publier**. Le projet web contient les fichiers d’application web du complément Office, et il s’agit donc du projet que vous publiez sur Azure.
    
3. Sur l’onglet **Publier** :

      - Choisissez **Microsoft Azure Application Service**.
      
      - Choisissez **Sélectionner**.

      - Choisissez **Publier**. 

6. Dans la boîte de dialogue **App Service**, recherchez et sélectionnez l’application web que vous avez créée à l’[étape 3 : Créer une application web dans Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure), puis cliquez sur **OK**. 

    Visual Studio publie le projet web pour votre complément Office sur votre site web Azure. Une fois le projet web publié par Visual Studio, votre navigateur s’ouvre et affiche une page web avec le texte « Votre application de service d’application a été créée. » Il s’agit de la page active par défaut pour l’application web.

7. Pour voir la page web pour votre complément, modifiez l’URL afin qu’elle utilise le protocole HTTPS et indiquez le chemin d’accès de la page HTML de votre complément (par exemple : https://VotreDomaine.azurewebsites.net/Home.html). Cela permet de confirmer que l’application web de votre complément est hébergée sur Azure. Copiez l’URL racine (par exemple : https://VotreDomaine.azurewebsites.net) ; vous en aurez besoin lorsque vous modifierez le fichier manifest de complément plus loin dans cet article.
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a>Étape 6 : Modifier et déployer le fichier manifeste XML

1. Dans Visual Studio avec l’exemple de complément Office ouvert dans l’**explorateur de solutions**, développez la solution pour que les deux projets s’affichent.
    
2. Développez le projet macro complémentaire Office (par exemple WordWebAddIn), le dossier manifest d’avec le bouton droit de la souris et sélectionnez **ouvrir**. Le fichier manifeste XML du complément s’ouvre.
    
3. Dans le fichier manifeste XML, recherchez et remplacez toutes les instances de « ~remoteAppUrl » par l’URL racine de l’application web du complément sur Azure. Il s’agit de l’URL que vous avez copiée précédemment une fois que vous avez publié l’application web du complément sur Azure (par exemple : https://VotreDomaine.azurewebsites.net). 
    
4. Choisissez **Fichier**, puis **Enregistrer tout**. Fermez le fichier manifeste XML du complément.
    
5. Retournez dans l’**explorateur de solutions**, cliquez avec le bouton droit de la souris sur le dossier du fichier manifeste et choisissez **Ouvrir le dossier dans l'Explorateur de fichiers**.
    
6. Copiez le fichier manifeste XML du complément (par exemple, WordWebAddIn.xml). 
    
7. Accédez au partage de fichiers réseau que vous avez créé à l’[étape 1 : Créer un dossier partagé](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) et collez le fichier manifeste dans le dossier.

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a>Étape 7 : Insérer et exécuter le complément dans l’application cliente Office

1. Démarrez Word 2016 et créez un document.
    
2. Sur le ruban, cliquez sur **Insérer** > **Mes compléments**. 
    
3. Dans la boîte de dialogue **Compléments Office**, choisissez **DOSSIER PARTAGÉ**. Word recherche le dossier que vous avez désigné comme catalogue de compléments approuvés (à l’[étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) et affiche les compléments dans la boîte de dialogue. Vous devriez voir l’icône de votre exemple de complément.
    
4. Cliquez sur l’icône de votre complément, puis choisissez **Ajouter**. Un bouton **Afficher le volet de tâches** pour votre complément est ajouté au ruban. 

5. Dans le ruban de l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches**. Le complément s’ouvre dans un volet de tâches à droite du document actif.
    
6. Vérifiez que le complément fonctionne en sélectionnant du texte dans le document et en choisissant le bouton **Mettre en surbrillance** dans le volet de tâches. 

## <a name="additional-resources"></a>Ressources supplémentaires

- [Publier votre complément Office](../publish/publish.md)
    
- [Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication](../publish/package-your-add-in-using-visual-studio.md)
    
