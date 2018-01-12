
# <a name="get-started-with-labsjs-for-office-mix"></a>Prise en main de LabsJS pour Office Mix



LabsJS comprend une API (labs.js), des exemples, de la documentation, ainsi que des fichiers associés que vous pouvez utiliser pour développer des ateliers interactifs, les intégrer à Office Mix, puis les afficher dans Microsoft PowerPoint. Ces ateliers sont, en fait, des Compléments Office que vous créez à l’aide de HTML5 et la bibliothèque JavaScript labs.js.

## <a name="labsjs-content"></a>Contenu de LabsJS

LabsJS fournit de la documentation, des exemples d’ateliers, ainsi que les fichiers requis pour créer ou publier vos propres ateliers pour Office Mix.


**Fichiers requis**


|**Fichier**|**Description**|
|:-----|:-----|
|labs-1.0.4.js|API JavaScript LabsJS pour le développement d’ateliers Office Mix. Ce fichier doit être inclus dans votre projet pour lui permettre d’intégrer Office Mix. Le fichier est également hébergé sur un réseau de distribution de contenu (CDN) à l’adresse  <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. Lorsque vous publiez votre application, vous devez l’associer au fichier sur le CDN.|
|labs-1.0.4.d.ts|Fichier de définition TypeScript pour labs.js. Il permet d’intégrer facilement votre code TypeScript à labs.js. Le fichier de définition fournit également une vue d’ensemble de tous les composants contenus dans labs.js. Vous pouvez télécharger TypeScript sur [http://www.typescriptlang.org/](http://www.typescriptlang.org/). Le fichier de définition a été créé pour la version 0.9.1.1 de TypeScript.|
|Historique|Historique des versions de la bibliothèque labs.js.|
|Labshost.html|Page web permettant de visualiser et de déboguer votre atelier par rapport à Office Mix, en dehors du contexte de PowerPoint. Pour utiliser cette page, saisissez votre URL dans la zone de texte principale et elle se charge dans le cadre. Les données échangées entre l’API et Office Mix lors de l’exécution dans PowerPoint, ou bien le lecteur de leçon Office Mix apparaît dans les zones de texte sur la droite. Les données peuvent également être pré-amorcées. Notez que les ateliers apparaissant dans la section Exemples représentent les Compléments Office Mix existantes en cours d’exécution dans le contexte hôte.|
|SampleManifest.xml|Exemple de manifeste d’Compléments Office à utiliser comme modèle pour créer votre propre manifeste d’application.|
|Simplelab.html|Exemple d’atelier créé à l’aide de labs.js. Il permet de sélectionner et d’insérer une page web afin de suivre l’utilisateur qui la visualise.|
|Simplelab.ts|Fichier TypeScript utilisé pour créer un exemple d’atelier.|
|Simplelab.js|Version JavaScript de l’exemple d’atelier. Ce fichier, comme simplelab.ts, utilisent l’API LabsJS.|

## <a name="set-up-your-development-environment"></a>Configuration de votre environnement de développement

La bibliothèque labs.js sert de couche d’abstraction sur la bibliothèque office.js (l’API d’Compléments Office), de sorte que les ateliers que vous créez à l’aide de la bibliothèque labs.js sont en fait des Compléments Office. Pour pouvoir utiliser la bibliothèque labs.js et exécuter ces ateliers dans Office Mix, vous devez d’abord vous définir en tant que développeur d’Compléments Office.


### <a name="register-for-an-office-365-developer-site"></a>Inscription auprès d’un site du développeur Office 365

La première étape consiste à vous inscrire auprès d’un Site du développeur Office 365. Cela vous permet d’héberger et de tester votre atelier avant de le soumettre à l’Office Store. Le site vous permet de publier votre complément sur Office Mix et de le tester dans un environnement réel.

Pour plus d’informations, voir [Configurer un environnement de développement pour les compléments pour SharePoint dans Office 365](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx). 


### <a name="set-up-an-app-catalog-on-sharepoint-online"></a>Configuration d’un catalogue d’applications sur SharePoint Online

Une fois que votre site du développeur est créé et mis en service, vous configurez ensuite un catalogue de complément sur SharePoint Online. Pour plus d’informations, voir [Configurer un catalogue de compléments dans Office 365](../../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

Pour Office Mix, vous utilisez un catalogue de complément afin de pouvoir insérer des compléments de pré-production dans une leçon et de réaliser des tests de bout en bout avant de soumettre les ateliers au magasin.


## <a name="create-your-lab"></a>Création de votre atelier

Pour créer votre premier atelier, suivez les étapes indiquées dans la [procédure pas à pas](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md). Elle explique comment créer un simple questionnaire vrai/faux. Voir [Procédure : Création de votre premier atelier pour Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md).


## <a name="publish-your-lab"></a>Publication de votre atelier

Après avoir créé votre atelier, vous pouvez le publier et le soumettre au magasin.


### <a name="create-and-upload-your-application-manifest"></a>Création et téléchargement de votre manifeste d’application

Le manifeste d’application est un document XML qui décrit votre atelier LabJS. Il indique l’URL où l’atelier est hébergé et fournit des détails sur celui-ci, y compris le nom d’affichage, la description, les icônes, la taille, etc.

Nous incluons un exemple de manifeste intitulé « SampleManifest.xml ». Pour obtenir plus d’informations sur le schéma de manifeste et un lien vers la définition du schéma, voir [Manifeste XML des compléments Office](../../../docs/overview/add-in-manifests.md).

Pour télécharger votre manifeste sur votre site SharePoint, accédez d’abord à votre catalogue d’applications, qui se trouve généralement à l’URL <code>https://\<your site\>/sites/AppCatalog</code>. Sélectionnez ensuite le bouton **Nouvelle application** et suivez les étapes pour télécharger votre manifeste d’application.


### <a name="update-your-powerpoint-2013-catalog"></a>Mise à jour de votre catalogue PowerPoint 2013

Ensuite, mettez à jour votre catalogue PowerPoint 2013. Après cela, vous pourrez vous connecter avec votre compte de développeur.

Commencez par mettre à jour votre catalogue PowerPoint 2013. Lancez PowerPoint 2013 et cliquez sur  **Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion de la confidentialité > Catalogues d’applications approuvés**. De là, ajoutez une référence à votre catalogue d’applications et choisissez  **OK**. PowerPoint 2013 vous demande de vous déconnecter pour que les modifications prennent effet. Déconnectez-vous.

Enfin, reconnectez-vous en utilisant votre compte du développeur. Choisissez votre nom d’ouverture de session dans le coin supérieur droit dans PowerPoint 2013 et connectez-vous en utilisant votre compte du développeur. Vous pouvez maintenant insérer votre complément.


### <a name="insert-publish-and-view-your-app"></a>Insertion, publication et visualisation de votre application

Pour insérer votre complément au catalogue, choisissez le ruban  **Insérer**, puis sélectionnez  **Stocker** dans la section **Applications**. Sélectionnez  **Mon organisation** et le complément apparaît dans votre catalogue de complément. Cliquez sur le complément, sélectionnez **Insérer** et votre complément (votre atelier) est inséré dans le document PowerPoint 2013.

Maintenant, vous pouvez profiter de toutes les fonctionnalités Office Mix disponibles pour publier votre leçon avec votre nouvel atelier.


 >**Important** :  Pour visualiser l’application, vous devez vous connecter à votre catalogue SharePoint avec le même navigateur dans lequel vous avez visualisé votre leçon. Seuls les utilisateurs authentifiés peuvent accéder aux catalogues SharePoint. C’est pourquoi vous devez d’abord vous connecter pour voir votre application. 


### <a name="submit-your-lab-to-the-office-store"></a>Soumission de votre application à l’Office Store

Pour soumettre votre atelier à l’Office Store, voir [Publier votre complément Office](../../publish/publish.md).


## <a name="additional-resources"></a>Ressources supplémentaires



- [Compléments Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Compléments Office](../../../docs/overview/office-add-ins.md)
    
- [Création de votre premier laboratoire pour Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)
    
