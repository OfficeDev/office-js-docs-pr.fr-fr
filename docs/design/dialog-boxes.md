# <a name="dialog-boxes-in-office-add-ins"></a>Boîtes de dialogue dans les compléments Office
 
Les boîtes de dialogue sont des surfaces qui flottent au-dessus de la fenêtre active de l’application Office. Vous pouvez utiliser les boîtes de dialogue afin de fournir un espace supplémentaire sur l’écran pour les tâches comme les pages de connexion impossibles à ouvrir directement dans un volet des tâches, ou pour les demandes de confirmation d’une action effectuée par un utilisateur, ou pour afficher des vidéos qui peuvent être trop petites si confinées à un volet des tâches.

**Exemple : Boîte de dialogue**

![Exemple d’image affichant une mise en page par défaut pour une boîte de dialogue](../../images/overview_withApp_dialog.png)

### <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:-----|:--------|
|<ul><li>Inclure un titre descriptif qui inclut le nom de votre complément, ainsi que la tâche en cours.</li></ul>|<ul><li>Ne pas ajouter le nom de votre société au titre.</li></ul>|
||<ul><li>Ne pas ouvrir une boîte de dialogue, sauf si le scénario l’exige.</li></ul>|

## <a name="implementation"></a>Implémentation

Pour voir un exemple relatif à l’implémentation d’une boîte de dialogue, consultez [Exemple d’interface API de boîte de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) dans GitHub.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Exemple de modèle UX](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [Ressources de développement GitHub](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Objet Dialogue](https://dev.office.com/reference/add-ins/shared/officeui.dialog)


