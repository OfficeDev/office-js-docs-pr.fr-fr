---
title: Compl?ments Office de contenu
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd0dcea7a3f37175a48946fc9dcd61d2b89f9c08
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="content-office-add-ins"></a>Compl?ments Office de contenu

Les compl?ments de contenu sont des surfaces qui peuvent ?tre incorpor?es directement dans des documents Word, Excel ou PowerPoint. Les compl?ments de contenu permettent aux utilisateurs d?acc?der aux contr?les d?interface qui ex?cutent le code pour modifier des documents ou afficher des donn?es d?une source de donn?es. Utilisez les compl?ments de contenu lorsque vous souhaitez incorporer des fonctionnalit?s directement dans le document.  

*Figure 1. Mise en page type pour les compl?ments de contenu*

![Exemple d?image affichant une mise en page typique pour des compl?ments de contenu.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Meilleures pratiques

- Inclure un ?l?ment de navigation ou de commande comme le CommandBar ou le tableau crois? dynamique en haut de votre compl?ment.
- Inclure un ?l?ment de la marque tel que le BrandBar en bas de votre compl?ment (s?applique aux compl?ments Word, Excel et PowerPoint uniquement).

## <a name="variants"></a>Variantes

Les tailles des compl?ments de contenu pour Word, Excel et PowerPoint dans le bureau Office 2016 et Office 365 sont sp?cifi?es par l?utilisateur.

## <a name="personality-menu"></a>Menu Caract?ristique

Les menus Caract?ristique peuvent entraver les ?l?ments de navigation et de commande se trouvant en haut ? droite du compl?ment. Voici les dimensions actuelles du menu Caract?ristique sur Windows et Mac.

Pour Windows, le menu Caract?ristique mesure 12 x 32 pixels, comme illustr?.

*Figure 2. Menu Caract?ristique sur Windows* 

![Image illustrant le menu Caract?ristique sur le bureau Windows](../images/personality-menu-win.png)


Pour Mac, le menu Caract?ristique mesure 26 x 26 pixels, mais flotte ? 8 pixels de la droite et ? 6 pixels du haut, ce qui permet d?augmenter l?espace occup? ? 34 x 32 pixels, comme illustr?.

*Figure 3. Menu Caract?ristique sur Mac*

![Image illustrant le menu Caract?ristique sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Impl?mentation

Pour consulter un exemple qui impl?mente un compl?ment de contenu, reportez-vous ? [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.

## <a name="support-considerations"></a>Consid?rations relatives ? la prise en charge
- V?rifiez si votre compl?ment Office fonctionne sur une [plateforme h?te Office sp?cifique](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability). 
- Certains compl?ments de contenu peuvent exiger que l?utilisateur accepte que le compl?ment lise et ?crive sur Excel ou PowerPoint. Vous pouvez d?clarer le [niveau des autorisations](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) que vous souhaitez attribuer ? votre utilisateur dans le manifeste du compl?ment.  
- Les compl?ments de contenu sont pris en charge dans Excel et PowerPoint dans Office 2013 et les versions ult?rieures. Si vous ouvrez un compl?ment dans une version d?Office qui ne prend pas en charge les compl?ments web Office, le compl?ment s?affichera comme une image.

## <a name="see-also"></a>Voir aussi
- [Disponibilit? des compl?ments Office sur les plateformes et les h?tes](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)
- [Office UI Fabric dans des compl?ments Office](https://docs.microsoft.com/en-us/office/dev/add-ins/design/office-ui-fabric) 
- [Mod?les de conception de l?exp?rience utilisateur pour les compl?ments Office](https://docs.microsoft.com/en-us/office/dev/add-ins/design/ux-design-patterns)
- [Demande d?autorisations d?utilisation de l?API dans des compl?ments de contenu et de volet des t?ches](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
