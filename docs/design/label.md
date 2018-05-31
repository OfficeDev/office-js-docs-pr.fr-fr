---
title: Composant d’étiquette dans la structure de l’interface utilisateur d’Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e9d6e9eaca918068b682725ee9236f6539641fa0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437240"
---
# <a name="label-component-in-office-ui-fabric"></a>Composant d’étiquette dans la structure de l’interface utilisateur d’Office

Utilisez des étiquettes pour nommer ou donner un titre à un composant ou un groupe de composants. Associées à un autre composant ou groupe de composants, les étiquettes doivent se trouver à proximité des composants ou des groupes associés. Certains composants ont des étiquettes prédéfinies comme les listes déroulantes ou les boutons bascule.
  
#### <a name="example-label-in-a-task-pane"></a>Exemple : Étiquette dans un volet de tâches

![Image illustrant l’étiquette](../images/overview-with-app-label.png)

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:------------|:--------------|
|Utilisez la casse pour une phrase, par exemple **Prénom**.|N’utilisez pas la casse pour un titre, par exemple **Prénom**.|
|Soyez court et concis.|N’utilisez pas de phrases complètes ni de signes de ponctuation complexes comme les virgules ou les points-virgules.|
|Lorsque vous ajoutez une étiquette à des composants, utilisez un nom ou une locution nominale courte comme texte d’étiquette.| |

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Étiquette par défaut**|À utiliser pour les étiquettes standard.|![Image de l’étiquette par défaut](../images/label.png)<br/>|
|**Étiquette désactivée**|À utiliser lorsque le composant associé est désactivé.|![Image d’étiquette désactivée](../images/label-disabled.png)<br/>|
|**Étiquette requise**|À utiliser lorsque le composant associé est requis.|![Image d’étiquette requise](../images/label-required.png)<br/>|

## <a name="implementation"></a>Implémentation

Pour plus d’informations, reportez-vous à [Étiquette](https://dev.office.com/fabric#/components/label) et [Démarrer avec un exemple de code React de la structure](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Voir aussi

- [Modèles de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric dans des compléments Office](office-ui-fabric.md)
