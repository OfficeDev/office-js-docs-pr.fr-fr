---
title: Office UI Fabric dans des compléments Office 
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: ec5b443d71e465a6e43cb0d871fc6fddd30e231a
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388037"
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric dans des compléments Office 

Office UI Fabric est une infrastructure frontale JavaScript permettant de créer des expériences pour Office et Office 365. Fabric propose des composants axés sur des visuels que vous pouvez étendre, retravailler et utiliser dans votre complément Office. Fabric utilisant le langage de création d’Office, ses composants d’expérience utilisateur ressemblent à une extension naturelle d’Office. 

Si vous créez un complément, nous vous encourageons à utiliser Office UI Fabric pour mettre au point l’expérience utilisateur. L’utilisation d’Office UI Fabric est facultative.

Les sections suivantes expliquent comment commencer à utiliser Fabric en fonction de vos besoins. 

## <a name="use-fabric-core-icons-fonts-colors"></a>Utiliser Fabric Core : icônes, polices, couleurs
Fabric Core contient les principaux éléments du langage de création tels que les icônes, les couleurs, le type et la grille.Fabric Core n’est pas dépendant de l’infrastructure. Fabric Core est utilisé par et inclus avec Fabric React.

Pour commencer à utiliser Fabric Core:

1. Ajoutez la référence CDN au code HTML sur votre page.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```   
    
2. Utilisez les polices et les icônes Fabric. 

    Pour utiliser une icône Fabric, incluez l’élément « i » sur votre page, puis référencez les classes appropriées. Vous pouvez contrôler la taille de l’icône en modifiant la taille de police. Par exemple, le code suivant montre comment créer une icône de tableau extra large qui utilise la couleur themePrimary (#0078d7). 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Pour rechercher des icônes supplémentaires disponibles dans Office UI Fabric, utilisez la fonctionnalité de recherche de la page [Icônes](https://developer.microsoft.com/fabric#/styles/icons). Lorsque vous trouvez une icône à utiliser dans votre complément, veillez à précéder le nom de l’icône de `ms-Icon--`. 

    Pour plus d’informations sur les tailles de police et les couleurs disponibles dans Office UI Fabric, voir [Typographie](https://developer.microsoft.com/fabric#/styles/typography) et [Couleurs](https://developer.microsoft.com/fabric#/styles/colors).
 
## <a name="use-fabric-components"></a>Utiliser les composants Fabric 
Fabric fournit une variété de composants UX que vous pouvez utiliser pour créer votre complément, y compris les types de composants suivants :

- Composants d’entrée- exemple, bouton, case à cocher et bouton bascule
- Composants de navigation- par exemple, tableau croisé dynamique, barre de navigation
- Composants de notification-par exemple, MessageBar et légende  

Pas tous les composants tissu sont recommandées pour les utiliser dans des compléments. Voici une liste des composants expérience utilisateur UX Fabric React recommandés pour les utiliser dans un complément:

- [Barre de navigation](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [Bouton](https://developer.microsoft.com/fabric#/components/button)
- [Case à cocher](https://developer.microsoft.com/fabric#/components/checkbox)
- [ChoiceGroup](https://developer.microsoft.com/fabric#/components/choicegroup)
- [Liste déroulante](https://developer.microsoft.com/fabric#/components/dropdown)
- [Étiquette](https://developer.microsoft.com/fabric#/components/label)
- [Liste](https://developer.microsoft.com/fabric#/components/list)
- [Tableau croisé dynamique](https://developer.microsoft.com/fabric#/components/pivot)
- [TextField](https://developer.microsoft.com/fabric#/components/textfield)
- [Bouton bascule](https://developer.microsoft.com/fabric#/components/toggle)

Vous pouvez utiliser différentes infrastructures JavaScript, comme Angular ou React, pour créer votre complément. Pour commencer à utiliser les composants Fabric avec votre infrastructure, consultez les ressources suivantes.

|**Infrastructure**|**Exemple**|
|:------------|:----------|
|**React**|[Utilisation d’Office UI Fabric React dans des compléments Office](using-office-ui-fabric-react.md )|
|**Angular**| Reportez-vous à [ngOfficeUIFabric](http://ngofficeuifabric.com/), qui est un projet communautaire avec des directives Angular 1.5, et [envisagez d’insérer des composants Fabric dans des composants Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components).|
