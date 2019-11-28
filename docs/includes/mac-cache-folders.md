Les compléments sont souvent mis en cache dans Office pour Mac, pour des raisons de performances. En règle générale, vous pouvez effacer le cache en rechargeant le complément. S’il existe plusieurs compléments dans le même document, le processus de suppression automatique du cache lors du rechargement n’est peut-être pas fiable.

Vous pouvez effacer le cache à l’aide du menu personnalité de n’importe quel complément du volet Office.
- Sélectionnez le menu personnalité. Ensuite, sélectionnez **Vider le cache Web**.
    > [!NOTE]
    > Vous devez exécuter macOS version 10.13.6 ou version ultérieure pour afficher le menu personnalité.
    
    ![Capture d’écran de l’option effacer le cache Web du menu personnalité.](../images/mac-clear-cache-menu.png)

Vous pouvez également effacer le cache manuellement en supprimant le contenu du `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` dossier.

> [!NOTE]
> Si ce dossier n’existe pas, recherchez les dossiers suivants et, le cas échéant, supprimez le contenu du dossier :
>    - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`où `{host}` se trouve l’hôte Office (par exemple `Excel`,)
>    - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`où `{host}` se trouve l’hôte Office (par exemple `Excel`,)
>    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
