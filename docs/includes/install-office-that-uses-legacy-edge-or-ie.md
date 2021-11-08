Utilisez la procédure suivante pour installer une version de l’Office d’abonnement qui utilise le Version antérieure de Microsoft Edge webview (EdgeHTML) pour exécuter des applications ou une version qui utilise Internet Explorer (Trident).

1. Dans n’Office application, **ouvrez** l’onglet Fichier sur le **ruban,** puis sélectionnez Office compte ou **compte.** Sélectionnez le bouton À propos du nom **_d’hôte_** (par exemple, **À propos de Word).**
1. Dans la boîte de dialogue qui s’ouvre, recherchez le numéro de build complet xx.x.xxxxx.xxxxx et faites-en une copie quelque part.
1. Téléchargez et installez [l’Office Deployment Tool](https://www.microsoft.com/download/details.aspx?id=49117).
1. Dans le dossier où vous avez installé l’outil (où se trouve le fichier), créez un fichier texte avec le nom et `setup.exe` `config.xml` ajoutez le contenu suivant.

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

    - Pour installer une version qui utilise l’ancienne version de Edge, changez-la en `16.0.11929.20946` .
    - Pour installer une version qui utilise Internet Explorer, changez-la en `16.0.10730.20348` .

1. Éventuellement, modifiez la valeur de pour installer les Office `OfficeClientEdition` 32 bits et modifiez la valeur selon les besoins pour installer Office dans `"32"` `Language ID` une autre langue.
1. Ouvrez une invite de commandes *en tant qu’administrateur.*
1. Accédez au dossier avec les `setup.exe` fichiers `config.xml` et les fichiers.
1. Exécutez la commande suivante.

    ```command&nbsp;line
    setup.exe /configure config.xml
    ```

    Cette commande installe Office. Le processus peut prendre plusieurs minutes.

1. [Clear the Office cache](../testing/clear-cache.md).

> [!IMPORTANT]
> Après l’installation, veillez à désactiver la mise à jour automatique de Office afin que Office ne soit pas mis à jour vers une version qui n’utilise pas le webview que vous souhaitez utiliser avant d’avoir terminé de l’utiliser. **Cela peut se produire en quelques minutes après l’installation.** Procédez comme suit.
>
> 1. Démarrez n Office application et ouvrez un nouveau document.
> 1. Ouvrez **l’onglet** Fichier sur le ruban, puis sélectionnez **Office compte** ou **compte.**
> 1. Dans la colonne **Informations sur le** produit, **sélectionnez Options de** mise à jour, puis **désactivez les mises à jour.** Si cette option n’est pas disponible, Office est déjà configuré pour ne pas se mettre à jour automatiquement.

Lorsque vous avez terminé d’utiliser l’ancienne version de Office, réinstallez votre version la plus récente en modifiant le fichier et en modifiant le numéro de build que vous avez copié `config.xml` `Version` précédemment. Répétez ensuite `setup.exe /configure config.xml` la commande dans une invite de commandes d’administrateur. Vous avez la possibilité de ré-activer les mises à jour automatiques.
