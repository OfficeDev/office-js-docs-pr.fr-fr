---
title: Compléments Office de contenu
description: ''
ms.date: 12/04/2017
---



# <a name="content-office-add-ins"></a>Compléments Office de contenu

Les compléments de contenu sont des surfaces qui peuvent être incorporées directement dans des documents Word, Excel ou PowerPoint. Les compléments de contenu permettent aux utilisateurs d’accéder aux contrôles d’interface qui exécutent le code pour modifier des documents ou afficher des données d’une source de données. Utilisez les compléments de contenu lorsque vous souhaitez incorporer des fonctionnalités directement dans le document.  

*Figure 1. Mise en page type pour les compléments de contenu*

![Exemple d’image affichant une mise en page typique pour des compléments de contenu.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Meilleures pratiques

- Inclure un élément de navigation ou de commande comme le CommandBar ou le tableau croisé dynamique en haut de votre complément.
- Inclure un élément de la marque tel que le BrandBar en bas de votre complément (s’applique aux compléments Word, Excel et PowerPoint uniquement).

## <a name="variants"></a>Variantes

Les tailles des compléments de contenu pour Word, Excel et PowerPoint dans le bureau Office 2016 et Office 365 sont spécifiées par l’utilisateur.

## <a name="personality-menu"></a>Menu Caractéristique

Les menus Caractéristique peuvent entraver les éléments de navigation et de commande se trouvant en haut à droite du complément. Voici les dimensions actuelles du menu Caractéristique sur Windows et Mac.

Pour Windows, le menu Caractéristique mesure 12 x 32 pixels, comme illustré.

*Figure 2. Menu Caractéristique sur Windows* 

![Image illustrant le menu Caractéristique sur le bureau Windows](../images/personality-menu-win.png)


Pour Mac, le menu Caractéristique mesure 26 x 26 pixels, mais flotte à 8 pixels de la droite et à 6 pixels du haut, ce qui permet d’augmenter l’espace occupé à 34 x 32 pixels, comme illustré.

*Figure 3. Menu Caractéristique sur Mac*

![Image illustrant le menu Caractéristique sur le bureau Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implémentation

Pour consulter un exemple qui implémente un complément de contenu, reportez-vous à [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) dans GitHub.

## <a name="see-also"></a>Voir aussi

- [Office UI Fabric dans des compléments Office](office-ui-fabric.md) 
- [Modèles de conception de l’expérience utilisateur pour les compléments Office](ux-design-patterns.md)
