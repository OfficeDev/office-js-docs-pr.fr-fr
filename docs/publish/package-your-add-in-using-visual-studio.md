# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication

Votre package de complément Office contient un [fichier manifeste](../overview/add-in-manifests.md) XML que vous allez utiliser pour publier le complément. Vous devez publier les fichiers d’application web de votre projet séparément. Cet article décrit le déploiement de votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2015.

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>Déploiement de votre projet web à l’aide de Visual Studio 2015

Procédez comme suit pour déployer votre projet web à l’aide de Visual Studio 2015.

1. Dans l’**explorateur de solutions**, ouvrez le menu contextuel du projet de complément, puis sélectionnez **Publier**.
    
    La page **Publier votre complément** s’ouvre.
    
2. Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau…** pour créer un profil.
    
     >**Remarque :**  Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.

    Si vous choisissez **Nouveau...**, l’Assistant **Créer un profil de publication** s’ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.
    
    Pour plus d’informations sur l’importation et la création de profils de publication, voir [Création d’un profil de publication](http://msdn.microsoft.com/fr-fr/library/dd465337.aspx#creating_a_profile).
    
3. Sur la page **Publier votre complément**, cliquez sur le lien **Déployer votre projet Web**.
    
    La boîte de dialogue **Publier le site web** s’affiche. Pour plus d’informations sur l’utilisation de cet Assistant, reportez-vous à l’article relatif à la [procédure de déploiement d’un projet web à l’aide de la publication en un clic dans Visual Studio](http://msdn.microsoft.com/fr-fr/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>Création d’un package de votre complément avec Visual Studio 2015

Procédez comme suit pour créer un package de votre projet de complément à l’aide de Visual Studio 2015.

1. Sur la page **Publier votre complément**, cliquez sur le lien **Empaqueter le complément**.
    
    L’Assistant **Publication des compléments SharePoint et Office** apparaît.
    
2. Dans la liste déroulante **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur **Terminer**.
    
    Vous devez spécifier une adresse qui commence par le préfixe HTTPS pour l’Assistant. L’utilisation d’un point de terminaison HTTPS pour votre site web est généralement recommandée, mais cela n’est pas obligatoire si vous ne comptez pas publier votre complément sur l’Office Store. Si vous souhaitez utiliser un point de terminaison HTTP pour votre site web, vous pouvez ouvrir le fichier manifeste XML dans un éditeur de texte une fois que le package a été créé et remplacer le préfixe HTTPS de votre site web par un préfixe HTTP. Pour plus d’informations, reportez-vous à [Pourquoi mes applications et compléments doivent-ils être sécurisés par une protection SSL ?](http://msdn.microsoft.com/fr-fr/library/jj591603#bk_q7).
    
     >**Remarque :**  Les sites web Azure fournissent automatiquement un point de terminaison HTTPS.

    Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication. 
    
Si vous prévoyez de soumettre votre complément à l’Office Store, vous pouvez cliquer sur le lien **Effectuer un test de validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté. Vous devez corriger tous les problèmes avant d’envoyer votre complément au Store.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Vous trouverez le fichier manifeste XML dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>Ressources supplémentaires



- [Publier votre complément Office](../publish/publish.md)
    
- [Soumission des compléments SharePoint et Office, ainsi que des applications web Office 365 dans l’Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
