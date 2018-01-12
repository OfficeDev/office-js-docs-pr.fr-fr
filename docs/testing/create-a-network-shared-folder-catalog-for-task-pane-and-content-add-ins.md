
# <a name="sideload-office-add-ins-for-testing"></a>Chargement de version test des compléments Office

Vous pouvez installer un complément Office à des fins de test dans un client Office s’exécutant sur Windows à l’aide d’un catalogue de dossiers partagés pour publier le manifeste sur un partage de fichiers réseau. 

Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour charger la version test de votre complément :

- [Chargement de version test des compléments Office dans Office Online](sideload-office-add-ins-for-testing.md)
- [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook](sideload-outlook-add-ins-for-testing.md)

La vidéo suivante présente la procédure de chargement de version test de votre complément dans la version de bureau Office ou Office Online.

<iframe width="560" height="315" src="https://www.youtube.com/embed/XXsAw2UUiQo" frameborder="0" allowfullscreen></iframe>


## <a name="share-a-folder"></a>Partager un dossier

1. Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.

2. Ouvrez le menu contextuel du dossier (clic droit), puis choisissez **Propriétés**.

3. Ouvrez l’onglet **Partage**.

4. Dans la page **Choisir les utilisateurs...**, ajoutez votre nom et celui des utilisateurs avec lesquels vous souhaitez partager votre complément. S’ils sont tous membres d’un groupe de sécurité, vous pouvez ajouter le groupe. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier. 

5. Choisissez **Partager** > **Terminer** > **Fermer**.

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>Spécifier le dossier partagé en tant que catalogue approuvé

      
3. Ouvrez un nouveau document dans Excel, Word ou PowerPoint.
    
4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
5. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
    
6. Choisissez **Catalogues de compléments approuvés**.
    
7. Dans la zone **URL du catalogue**, entrez le chemin d’accès réseau complet au catalogue de dossiers partagés, puis choisissez **Ajouter un catalogue**.
    
8. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.

9. Fermez l’application Office afin que vos modifications prennent effet.
    
## <a name="sideload-your-add-in"></a>Charger votre complément


1. Placez le fichier manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.

    >**Important :**  pour contribuer à sécuriser les compléments accédant à des services et données externes, votre complément doit utiliser un protocole sécurisé tel que HTTPS pour se connecter aux services et données externes. Vous devez utiliser HTTPS si votre complément utilise des commandes de complément.

2. Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.

3. Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.

4. Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.


## <a name="additional-resources"></a>Ressources supplémentaires

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)
    
