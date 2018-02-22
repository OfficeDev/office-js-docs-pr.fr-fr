---
title: Composant de case à cocher dans Office UI Fabric
description: ''
ms.date: 12/04/2017
---

# <a name="checkbox-component-in-office-ui-fabric"></a>Composant de case à cocher dans Office UI Fabric

Une case à cocher est un élément de l’interface utilisateur qui permet aux utilisateurs d’activer ou de désactiver des options dans les compléments. Utilisez les cases à cocher pour permettre aux utilisateurs de sélectionner des options. En outre, une case à cocher peut être associée à un contrôle. Lorsque la case à cocher est sélectionnée ou désélectionnée, le comportement du contrôle associé change. Par exemple, le contrôle associé peut basculer entre l’état visible ou masqué.
  
#### <a name="example-check-box-in-a-task-pane"></a>Exemple : Case à cocher dans un volet des tâches

![Image illustrant une case à cocher](../images/overview-with-app-checkbox.png)

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:------------|:--------------|
|Utiliser les cases à cocher pour indiquer l’état.<br/><br/>![À faire : exemple de case à cocher](../images/checkbox-do.png)<br/>|Ne pas utiliser les cases à cocher pour afficher/indiquer une action.<br/><br/>![À ne pas faire : exemple de case à cocher](../images/checkbox-dont.png)<br/>|
|Utiliser plusieurs cases à cocher lorsque les utilisateurs peuvent sélectionner plusieurs options qui ne s’excluent pas mutuellement.|Ne pas utiliser de case à cocher lorsque les utilisateurs ne peuvent choisir qu’une seule option. Utiliser les cases d’option lorsqu’ils ne doivent sélectionner qu’une seule option.|
|Autoriser les utilisateurs à choisir n’importe quelle combinaison d’options lorsque plusieurs cases à cocher sont regroupées.|Ne pas placer deux groupes de cases à cocher l’un à côté de l’autre. Séparer les deux groupes avec des étiquettes.|
|Utiliser une case à cocher unique pour un paramètre secondaire. Par exemple, la case à cocher **Mémoriser mes informations** est un paramètre secondaire utilisé dans un scénario de connexion.|Ne pas utiliser de cases à cocher pour activer et désactiver des paramètres. Pour passer d’un état activé à désactivé ou vice-versa, utiliser un bouton bascule.|

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Case à cocher non contrôlée**|À utiliser comme état de case à cocher par défaut. |![Image Case à cocher non contrôlée](../images/checkbox-unchecked.png)|
|**Case à cocher non contrôlée avec la valeur Vrai sélectionnée par défaut**|À utiliser lorsque l’instance de case à cocher conserve son propre état |![Image Case à cocher non contrôlée avec la valeur Vrai sélectionnée par défaut](../images/checkbox-checked.png)|
|**Case à cocher non contrôlée désélectionnée avec la valeur Vrai sélectionnée par défaut**|État désactivé de la case à cocher. |![Image Case à cocher non contrôlée désélectionnée avec la valeur Vrai sélectionnée par défaut](../images/checkbox-disabled.png)|
|**Case à cocher contrôlée**|L’état sélectionné de cette case à cocher est décidé à un autre endroit de votre interface utilisateur. Dans ce scénario, la valeur correcte est transmise à la case à cocher par un événement **onChange** et le nouveau rendu de l’interface utilisateur. |![Case à cocher contrôlée](../images/checkbox-unchecked.png)|

## <a name="implementation"></a>Implémentation

Pour plus d’informations, reportez-vous à [Case à cocher](https://dev.office.com/fabric#/components/checkbox) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Voir aussi

- [Modèles de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric dans des compléments Office](office-ui-fabric.md)
