# <a name="textfield-component-in-office-ui-fabric"></a>Composant TextField dans Office UI Fabric

Un champ de texte permet aux utilisateurs d’entrer du texte. Il est généralement utilisé pour capturer une seule ligne de texte mais peut être configuré pour capturer plusieurs lignes de texte. Le texte s’affiche à l’écran dans un format simple et uniforme.
  
#### <a name="example-textfield-in-a-task-pane"></a>Exemple : TextField dans un volet des tâches

![Image illustrant le composant TextField](../../images/overview_withApp_textField.png)

<br/>

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:------------|:--------------|
|Utiliser des champs de texte pour accepter la saisie de données sur un formulaire ou une page.|Ne pas utiliser de champs de texte pour rendre une copie de base dans un élément de corps d’une page.|
|Étiqueter les champs de texte avec des noms utiles.|Ne pas utiliser de champs de texte pour saisir une date ou une heure. Utiliser plutôt un sélecteur de date et heure.|
|Utiliser un texte de l’espace réservé concis pour spécifier le contenu qui doit être saisi.|Ne pas utiliser de champs de texte si des options d’entrée valides peuvent être prédéfinies. Utiliser plutôt une liste déroulante.|
|Fournir tous les états appropriés pour les champs de texte (statique, pointage, focus, engagé, non disponible, erreur).||
|Marquer clairement les champs obligatoires et facultatifs.||
|Si possible, mettre en forme les champs de texte en fonction du format de données attendu. Par exemple, lors de la capture d’un numéro de téléphone à 10 chiffres, utiliser trois champs distincts pour stocker les différentes parties du numéro de téléphone.||

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Default TextField (champ de texte par défaut)**|À utiliser comme champ de texte par défaut.|![Image Default TextField (champ de texte par défaut) ](../../images/textfieldDefault.png)<br/>|
|**Disabled TextField (champ de texte désactivé)**|À utiliser lorsque le champ de texte est désactivé.|![Image Disabled TextField (champ de texte désactivé)](../../images/textfieldDisabled.png)<br/>|
|**Required TextField (champ de texte obligatoire)**|À utiliser lorsque le champ de texte est obligatoire.|![Image Required TextField (champ de texte obligatoire)](../../images/textfieldRequired.png)<br/>|
|**TextField with a placeholder (champ de texte avec un espace réservé)**|À utiliser lorsqu’un texte de l’espace réservé est nécessaire.|![Image TextField with a placeholder (champ de texte avec un espace réservé)](../../images/textfieldPlaceholder.png)<br/>|
|**TextField with multiple lines (champ de texte avec plusieurs lignes)**|À utiliser lorsque plusieurs lignes de texte sont nécessaires.|![Image TextField with a placeholder (champ de texte avec un espace réservé)](../../images/textfieldMulti.png)<br/>|

## <a name="implementation"></a>Implémentation

Pour plus d’informations, voir [TextField](https://dev.office.com/fabric#/components/textfield) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="additional-resources"></a>Ressources supplémentaires

- [Modèles de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office UI Fabric dans des compléments Office](office-ui-fabric.md)
