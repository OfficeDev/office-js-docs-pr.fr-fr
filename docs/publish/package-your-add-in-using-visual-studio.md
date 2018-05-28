---
title: Cr?er le package de votre compl?ment ? l?aide de Visual Studio pour pr?parer la publication
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e03959294536eeb416a1531d2d281ba83f2d3732
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Cr?er le package de votre compl?ment ? l?aide de Visual Studio pour pr?parer la publication

Votre package de compl?ment Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous allez utiliser pour publier le compl?ment. Vous devez publier les fichiers d?application web de votre projet s?par?ment. Cet article d?crit le d?ploiement de votre projet web et l?empaquetage de votre compl?ment ? l?aide de Visual Studio 2015.

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>D?ploiement de votre projet web ? l?aide de Visual Studio 2015

Proc?dez comme suit pour d?ployer votre projet web ? l?aide de Visual Studio 2015.

1. Dans l?**explorateur de solutions**, ouvrez le menu contextuel du projet de compl?ment, puis s?lectionnez **Publier**.
    
    La page **Publier votre compl?ment** s?ouvre.
    
2. Dans la liste d?roulante **Profil actuel**, s?lectionnez un profil ou choisissez **Nouveau?** pour cr?er un profil.
    
    > [!NOTE]
    > Un profil de publication indique le serveur sur lequel vous effectuez le d?ploiement, les informations d?identification n?cessaires pour se connecter au serveur, les bases de donn?es ? d?ployer, ainsi que d?autres options de d?ploiement.

    Si vous choisissez  **Nouveau...**, l?Assistant **Cr?er un profil de publication** s?ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication ? partir d?un site web d?h?bergement comme Microsoft Azure ou cr?er un profil et ajouter votre serveur, vos informations d?identification et d?autres param?tres, comme d?crit dans la proc?dure suivante.
    
    Pour plus d?informations sur l?importation et la cr?ation de profils de publication, voir [Cr?ation d?un profil de publication](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. Sur la page  **Publier votre compl?ment**, cliquez sur le lien  **D?ployer votre projet Web**.
    
    La bo?te de dialogue **Publier Web** appara?t. Pour plus d?information sur l?utilisation de cet assistant, reportez-vous ? l?article [Proc?dure?: D?ployer un projet d?application Web ? l?aide de la publication en un clic dans Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>Cr?ation d?un package de votre compl?ment avec Visual Studio 2015

Proc?dez comme suit pour cr?er un package de votre projet de compl?ment ? l?aide de Visual Studio 2015.

1. Sur la page **Publier votre compl?ment**, cliquez sur le lien **Empaqueter le compl?ment**.
    
    L?Assistant **Publication des compl?ments SharePoint et Office** appara?t.
    
2. Dans la liste d?roulante **O? votre site web est-il h?berg? ?**, s?lectionnez ou saisissez l?URL HTTPS du site web qui h?bergera les fichiers de contenu de votre compl?ment, puis cliquez sur **Terminer**. 
    
    Vous devez sp?cifier une URL qui commence par le pr?fixe HTTPS pour terminer cet assistant. Si vous souhaitez utiliser un point de terminaison HTTP pour votre site web, vous pouvez ouvrir le fichier manifeste XML dans un ?diteur de texte une fois que le package a ?t? cr?? et remplacer le pr?fixe HTTPS de votre site web par un pr?fixe HTTP. 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Les sites Web Azure fournissent automatiquement un point de terminaison HTTPS.

    Visual Studio g?n?re les fichiers n?cessaires ? la publication de votre compl?ment, puis ouvre le dossier de sortie de publication. 
    
Si vous pr?voyez de soumettre votre compl?ment ? AppSource, vous pouvez s?lectionner le lien **Effectuer la v?rification de la validation** pour identifier les probl?mes susceptibles d?emp?cher votre compl?ment d??tre accept?. Vous devez r?gler tous ces probl?mes avant de soumettre votre compl?ment au magasin.

Vous pouvez d?sormais t?l?charger votre manifeste XML ? l?emplacement appropri? pour [publier votre compl?ment](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>Voir aussi

- [Publier votre compl?ment Office](../publish/publish.md)
- [Mise ? disposition de vos solutions sur AppSource et dans Office](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)
    
