---
title: Office UI Fabric dans des compl?ments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8fafe8a68c477868c12bff61c7f9ff23fc7314e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="office-ui-fabric-in-office-add-ins"></a>Office UI Fabric dans des compl?ments Office 

Office UI Fabric est une infrastructure frontale JavaScript permettant de cr?er des exp?riences pour Office et Office 365. Fabric propose des composants ax?s sur des visuels que vous pouvez ?tendre, retravailler et utiliser dans votre compl?ment Office. Fabric utilisant le langage de cr?ation d?Office, ses composants d?exp?rience utilisateur ressemblent ? une extension naturelle d?Office. 

Si vous cr?ez un compl?ment, nous vous encourageons ? utiliser Office UI Fabric pour mettre au point l?exp?rience utilisateur. L?utilisation d?Office UI Fabric est facultative.

Les sections suivantes expliquent comment commencer ? utiliser Fabric en fonction de vos besoins. 

## <a name="use-fabric-core-icons-fonts-colors"></a>Utiliser Fabric Core : ic?nes, polices, couleurs
Fabric Core contient les principaux ?l?ments du langage de cr?ation tels que les ic?nes, les couleurs, le type et la grille. Fabric Core n?est pas d?pendant de l?infrastructure. Les composants JS et React de la structure utilisent Fabric Core.

Pour commencer ? utiliser Fabric Core :

1. Ajoutez la r?f?rence CDN au code HTML sur votre page.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. Utilisez les polices et les ic?nes Fabric. 

    Pour utiliser une ic?ne Fabric, incluez l??l?ment ? i ? sur votre page, puis r?f?rencez les classes appropri?es. Vous pouvez contr?ler la taille de l?ic?ne en modifiant la taille de police. Par exemple, le code suivant montre comment cr?er une ic?ne de tableau extra large qui utilise la couleur themePrimary (#0078d7). 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Pour rechercher des ic?nes suppl?mentaires disponibles dans Office UI Fabric, utilisez la fonctionnalit? de recherche de la page [Ic?nes](https://dev.office.com/fabric#/styles/icons). Lorsque vous trouvez une ic?ne ? utiliser dans votre compl?ment, veillez ? pr?c?der le nom de l?ic?ne de `ms-Icon--`. 

    Pour plus d?informations sur les tailles de police et les couleurs disponibles dans Office UI Fabric, voir [Typographie](https://dev.office.com/fabric#/styles/typography) et [Couleurs](https://dev.office.com/fabric#/styles/colors).
 
## <a name="use-fabric-components"></a>Utiliser les composants Fabric 
Fabric fournit une vari?t? de composants UX que vous pouvez utiliser pour cr?er votre compl?ment, y compris les types de composants suivants :

- Composants d?entr?e - par exemple, bouton, case ? cocher et bouton bascule
- Composants de navigation - par exemple, tableau crois? dynamique, barre de navigation
- Composants de notification - par exemple, MessageBar et l?gende  

Il n?est pas recommand? d?utiliser tous les composants Fabric dans des compl?ments. Nous fournissons des conseils sur l?utilisation des composants recommand?s dans cette section. Par exemple, pour savoir comment utiliser un bouton Fabric dans votre compl?ment, voir [Bouton](button.md). 

Vous pouvez utiliser diff?rentes infrastructures JavaScript, comme Angular ou React, pour cr?er votre compl?ment. Pour commencer ? utiliser les composants Fabric avec votre infrastructure, consultez les ressources suivantes.

|**Infrastructure**|**Exemple**|
|:------------|:----------|
|**React**|[Utilisation d?Office UI Fabric React dans des compl?ments Office](using-office-ui-fabric-react.md )|
|**Angular**| Reportez-vous ? [ngOfficeUIFabric](http://ngofficeuifabric.com/), qui est un projet communautaire avec des directives Angular 1.5, et [envisagez d?ins?rer des composants Fabric dans des composants Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components).|
