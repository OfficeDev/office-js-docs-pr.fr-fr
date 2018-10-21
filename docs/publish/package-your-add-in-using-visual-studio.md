---
title: Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication | Microsoft Docs
description: Comment déployer votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2017.
ms.date: 01/25/2018
ms.openlocfilehash: 3515f88e41bc5f0af62a3b043beae5177f3291ac
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681762"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication

Votre package de complément Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous utiliserez pour publier le complément. Vous devez publier séparément les fichiers d’application web de votre projet. Cet article décrit le déploiement de votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2017.

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>Déploiement de votre projet web à l’aide de Visual Studio 2017

Procédez comme suit pour déployer votre projet web à l’aide de Visual Studio 2017.

1. Dans l’**explorateur de solutions**, ouvrez le menu contextuel du projet de complément, puis sélectionnez **Publier**.
    
    La page **Publier votre complément** s’ouvre.
    
2. Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau…** pour créer un profil.
    
    > [!NOTE]
    > Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.

    Si vous choisissez  **Nouveau...**, l’Assistant **Créer un profil de publication** s’ouvre. Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.
    
    Pour plus d’informations sur l’importation et la création de profils de publication, voir [Création d’un profil de publication](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).
    
3. Sur la page  **Publier votre complément**, cliquez sur le lien  **Déployer votre projet Web**.
    
    La boîte de dialogue **Publier** apparaît. Pour plus d’information sur l’utilisation de cet assistant, reportez-vous à l’article [Procédure : Déployer un projet d’application Web à l’aide de la publication en un clic dans Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>Création d’un package de votre complément avec Visual Studio 2017

Procédez comme suit pour créer un package de votre projet de complément à l’aide de Visual Studio 2017.

1. Sur la page **Publier votre complément**, cliquez sur le bouton **Empaqueter le complément**.
    
    Un Assistant s’affiche avec la page **Empaqueter le complément**.
    
2. Dans la liste déroulante  **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur  **Terminer**.
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Les sites web Azure fournissent automatiquement un point de terminaison HTTPS.

    Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication.
    
Si vous prévoyez d’envoyer votre complément à AppSource, vous pouvez sélectionner le bouton **Effectuer la vérification de la validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté. Vous devez régler tous ces problèmes avant d’envoyer votre complément au magasin.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>Voir aussi

- [Publier votre complément Office](../publish/publish.md)
- [Mise à disposition de vos solutions sur AppSource et dans Office](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
