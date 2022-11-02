Les compléments sont souvent mis en cache dans Office sur Mac pour des raisons de performances. En règle générale, vous pouvez effacer le cache en rechargeant le complément. En présence de plusieurs compléments dans le même document, il se peut que le processus d’effacement automatique du cache lors du rechargement ne fonctionne pas systématiquement.

### <a name="use-the-personality-menu-to-clear-the-cache"></a>Utiliser le menu personnalité pour effacer le cache

Vous pouvez vider le cache à l’aide du menu personnalité de n’importe quel complément du volet Office. Toutefois, étant donné que le menu de personnalité n’est pas pris en charge dans les compléments Outlook, vous pouvez essayer [d’effacer le cache manuellement](#clear-the-cache-manually) si vous utilisez Outlook.

- Sélectionnez le menu personnalité. Sélectionnez **Effacer le cache web**.
    > [!NOTE]
    > Vous devez exécuter macOS version 10.13.6 ou ultérieure pour afficher le menu personnalité.

    ![Capture d'écran de l'option « Effacer le cache Web » dans le menu « Personnalité ».](../images/mac-clear-cache-menu.png)

### <a name="clear-the-cache-manually"></a>Effacer le cache manuellement

Vous pouvez également effacer le cache manuellement en supprimant le contenu du dossier `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`. Recherchez ce dossier via le terminal.

> [!NOTE]
> Si ce dossier n’existe pas, recherchez les dossiers suivants via le terminal et, le cas échéant, supprimez le contenu du dossier.
>
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` où `{host}` est l’application Office (par exemple, `Excel`)
> - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` où `{host}` est l’application Office (par exemple, `Excel`)
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
> - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`
>
> Pour rechercher ces dossiers via le Finder, vous devez définir finder pour afficher les fichiers masqués. Finder affiche les dossiers dans le répertoire **Conteneurs** par nom de produit, par exemple **Microsoft Excel** au lieu de **com.microsoft.Excel**.