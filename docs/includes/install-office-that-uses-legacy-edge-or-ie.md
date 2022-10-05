Utilisez la procédure suivante pour installer une version d’Office d’abonnement Microsoft 365 qui utilise la vue web Version antérieure de Microsoft Edge (EdgeHTML) pour exécuter des compléments ou une version qui utilise Internet Explorer (Trident).

1. Dans n’importe quelle application Office, ouvrez l’onglet **Fichier** sur le ruban, puis sélectionnez **Compte Office** ou **Compte**. Sélectionnez le bouton **À propos _du nom d’hôte_** (par exemple, **À propos de Word**).
1. Dans la boîte de dialogue qui s’ouvre, recherchez le numéro de build xx.x.xxxxx.xxxxx complet et faites-en une copie quelque part.
1. Télécharger [l’outil Déploiement d’Office](https://www.microsoft.com/download/details.aspx?id=49117).
1. Exécutez le fichier téléchargé pour extraire l’outil. Vous êtes invité à choisir l’emplacement d’installation de l’outil.
1. Dans le dossier où vous avez installé l’outil (où se trouve le `setup.exe` fichier), créez un fichier texte portant le nom `config.xml` et ajoutez le contenu suivant.

    ```xml
    <Configuration>
      <Add OfficeClientEdition="64" Channel="SemiAnnual" Version="16.0.xxxxx.xxxxx">
        <Product ID="O365ProPlusRetail">
          <Language ID="en-us" />
        </Product>
      </Add>
    </Configuration>
    ```

1. Modifiez la `Version` valeur.

    - Pour installer une version qui utilise Edge Legacy, remplacez-la `16.0.11929.20946`par .
    - Pour installer une version qui utilise Internet Explorer, remplacez-la `16.0.10730.20348`par .

1. Si vous le souhaitez, modifiez la valeur pour installer Office 32 bits et modifiez la `Language ID` valeur en fonction des `OfficeClientEdition` `"32"` besoins pour installer Office dans une autre langue.
1. Ouvrez une invite de commandes *en tant qu’administrateur*.
1. Accédez au dossier contenant les fichiers et `config.xml` les `setup.exe` fichiers.
1. Exécutez la commande suivante :

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    Cette commande installe Office. Le processus peut prendre plusieurs minutes.

1. [Effacez le cache Office](../testing/clear-cache.md).

> [!IMPORTANT]
> Après l’installation, veillez à désactiver la mise à jour automatique d’Office afin qu’Office ne soit pas mis à jour vers une version qui n’utilise pas la vue web que vous souhaitez utiliser avant d’avoir terminé de l’utiliser. **Cela peut se produire dans les minutes qui suivent l’installation.** Procédez comme suit.
>
> 1. Démarrez une application Office et ouvrez un nouveau document.
> 1. Ouvrez l’onglet **Fichier** dans le ruban, puis sélectionnez **Compte Office** ou **Compte**.
> 1. Dans la colonne **Informations sur le produit**, sélectionnez **Options de mise à jour**, puis **Désactivez Mises à jour**. Si cette option n’est pas disponible, Office est déjà configuré pour ne pas être mis à jour automatiquement.

Lorsque vous avez terminé d’utiliser l’ancienne version d’Office, réinstallez votre version la plus récente en modifiant le `config.xml` fichier et en remplaçant le `Version` numéro de build que vous avez copié précédemment. Répétez ensuite la `setup.exe /configure config.xml` commande dans une invite de commandes administrateur. Si vous le souhaitez, réactivez les mises à jour automatiques.
