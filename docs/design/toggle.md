# <a name="toggle-component-in-office-ui-fabric"></a>Composant de bouton bascule dans Office UI Fabric

Les boutons bascule représentent un commutateur physique permettant d’activer ou de désactiver des éléments. Utilisez les boutons bascule pour présenter deux options qui s’excluent mutuellement (par exemple, activé et désactivé) lorsque le choix d’une option provoque une action immédiate.
  
#### <a name="example-toggle-in-a-task-pane"></a>Exemple : Bouton bascule dans un volet des tâches


![Image illustrant le composant de bouton bascule](../images/overview_withApp_toggle.png)

<br/>

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:------------|:--------------|
|Utiliser les boutons bascule pour les paramètres binaires lorsque les modifications sont immédiatement appliquées.<br/><br/>![Exemple de bouton bascule À faire](../images/toggleDo.png)<br/>|Ne pas utiliser de boutons bascule si les utilisateurs doivent effectuer une étape supplémentaire avant que les modifications prennent effet.<br/><br/>![Exemple de bouton bascule À ne pas faire](../images/toggleDont.png)<br/>|
|Ne remplacer les étiquettes **On** et **Off** que s’il existe des étiquettes plus spécifiques à utiliser pour un paramètre. Utiliser des étiquettes courtes (3 à 4 caractères) qui représentent des opposés binaires.| |

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Enabled and checked (Activé et sélectionné)**|À utiliser lorsque l’état basculé est actif.|![Image Enabled and checked (Activé et sélectionné)](../images/toggleEnabledOn.png)<br/>|
|**Enabled and unchecked (Activé et désélectionné)**|À utiliser lorsque l’état basculé est inactif.|![Image Enabled and unchecked (Activé et désélectionné)](../images/toggleEnabledOff.png)<br/>|
|**Disabled and checked (Désactivé et sélectionné)**|À utiliser lorsque l’état actif ne peut pas être modifié.|![Image Disabled and checked (Désactivé et sélectionné)](../images/toggleDisabledOn.png)<br/>|
|**Disabled and unchecked (Désactivé et désélectionné)**|À utiliser lorsque l’état inactif ne peut pas être modifié.|![Image Disabled and unchecked (Désactivé et désélectionné)](../images/toggleDisabledOff.png)<br/>|

## <a name="implementation"></a>Implémentation

Pour plus d’informations, voir [Bouton bascule](https://dev.office.com/fabric#/components/toggle) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="additional-resources"></a>Ressources supplémentaires

- [Modèles de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office UI Fabric dans des compléments Office](office-ui-fabric.md)
