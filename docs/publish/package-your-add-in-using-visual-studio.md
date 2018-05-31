---
title: Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e03959294536eeb416a1531d2d281ba83f2d3732
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438752"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication

Votre package de complément Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous allez utiliser pour publier le complément. Vous devez publier les fichiers d’application web de votre projet séparément. Cet article décrit le déploiement de votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2015.

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>Déploiement de votre projet web à l’aide de Visual Studio 2015

Procédez comme suit pour déployer votre projet web à l’aide de Visual Studio 2015.

1. Dans l’**explorateur de solutions**, ouvrez le menu contextuel du projet de complément, puis sélectionnez **Publier**.
    
    La page **Publier votre complément** s’ouvre.
    
2. Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau…** pour créer un profil.
    
    > [!NOTE]
    > Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.

    Si vous choisissez  **Nouveau...**, l’Assistant **Créer un profil de publication** s’ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.
    
    Pour plus d’informations sur l’importation et la création de profils de publication, voir [Création d’un profil de publication](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile).
    
3. Sur la page  **Publier votre complément**, cliquez sur le lien  **Déployer votre projet Web**.
    
    La boîte de dialogue **Publier Web** apparaît. Pour plus d’information sur l’utilisation de cet assistant, reportez-vous à l’article [Procédure : Déployer un projet d’application Web à l’aide de la publication en un clic dans Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>Création d’un package de votre complément avec Visual Studio 2015

Procédez comme suit pour créer un package de votre projet de complément à l’aide de Visual Studio 2015.

1. Sur la page **Publier votre complément**, cliquez sur le lien **Empaqueter le complément**.
    
    L’Assistant **Publication des compléments SharePoint et Office** apparaît.
    
2. Dans la liste déroulante **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL HTTPS du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur **Terminer**. 
    
    Vous devez spécifier une URL qui commence par le préfixe HTTPS pour terminer cet assistant. Si vous souhaitez utiliser un point de terminaison HTTP pour votre site web, vous pouvez ouvrir le fichier manifeste XML dans un éditeur de texte une fois que le package a été créé et remplacer le préfixe HTTPS de votre site web par un préfixe HTTP. 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Les sites Web Azure fournissent automatiquement un point de terminaison HTTPS.

    Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication. 
    
Si vous prévoyez de soumettre votre complément à AppSource, vous pouvez sélectionner le lien **Effectuer la vérification de la validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté. Vous devez régler tous ces problèmes avant de soumettre votre complément au magasin.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>Voir aussi

- [Publier votre complément Office](../publish/publish.md)
- [Mise à disposition de vos solutions sur AppSource et dans Office](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)
    
