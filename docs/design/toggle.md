---
title: Composant de bouton bascule dans Office UI Fabric
description: ''
ms.date: 12/04/2017
---

# <a name="toggle-component-in-office-ui-fabric"></a>Composant de bouton bascule dans Office UI Fabric

Les boutons bascules sont des commutateurs physiques qui activent ou désactivent des éléments. Utilisez les boutons bascules pour présenter deux options qui s’excluent mutuellement (par exemple, on et off), lorsque le choix d’une option provoque une action immédiate.
  
#### <a name="example-toggle-in-a-task-pane"></a>Exemple : bouton bascule dans un volet Office

![Image illustrant le composant de bouton bascule](../images/overview-with-app-toggle.png)

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:------------|:--------------|
|Utiliser les boutons bascule pour les paramètres binaires lorsque les modifications sont immédiatement appliquées.<br/><br/>![Exemple de bouton bascule À faire](../images/toggle-do.png)<br/>|Ne pas utiliser de boutons bascule si les utilisateurs doivent effectuer une étape supplémentaire avant que les modifications prennent effet.<br/><br/>![Exemple de bouton bascule À ne pas faire](../images/toggle-dont.png)<br/>|
|Remplacer les étiquettes **On** et **Off** uniquement s’il existe des étiquettes plus spécifiques à utiliser pour un paramètre. Utiliser des étiquettes courtes (3 à 4 caractères) qui représentent des opposés binaires.| |

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Enabled and checked (Activé et sélectionné)**|À utiliser lorsque l’état basculé est actif.|![Image Enabled and checked (Activé et sélectionné)](../images/toggle-enabled-on.png)<br/>|
|**Enabled and unchecked (Activé et désélectionné)**|À utiliser lorsque l’état basculé est inactif.|![Image Enabled and unchecked (Activé et désélectionné)](../images/toggle-enabled-off.png)<br/>|
|**Disabled and checked (Désactivé et sélectionné)**|À utiliser lorsque l’état actif ne peut pas être modifié.|![Image Disabled and checked (Désactivé et sélectionné)](../images/toggle-disabled-on.png)<br/>|
|**Disabled and unchecked (Désactivé et désélectionné)**|À utiliser lorsque l’état inactif ne peut pas être modifié.|![Image Disabled and unchecked (Désactivé et désélectionné)](../images/toggle-disabled-off.png)<br/>|

## <a name="implementation"></a>Implémentation

Pour plus d’informations, reportez-vous à [Bouton bascule](https://dev.office.com/fabric#/components/toggle) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Voir aussi

- [Modèles de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric dans des compléments Office](office-ui-fabric.md)
