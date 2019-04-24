---
title: Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication | Microsoft Docs
description: Déploiement de votre projet web et empaquetage de votre complément à l’aide de Visual Studio 2017.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 9233ebed217c9e4cc5def0dace67043f29462296
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451086"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication

Votre package de complément Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous allez utiliser pour publier le complément. Vous devez publier les fichiers d’application web de votre projet séparément. Cet article décrit le déploiement de votre projet web et l’empaquetage de votre complément à l’aide de Visual Studio 2017.

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>Déploiement de votre projet web à l’aide de Visual Studio 2017

Procédez comme suit pour déployer votre projet web à l’aide de Visual Studio 2017.

1. Dans l’**explorateur de solutions**, ouvrez le menu contextuel du projet de complément, puis sélectionnez **Publier**.

    La page **Publier votre complément** s’ouvre.

2. Dans la liste déroulante **Profil actuel**, sélectionnez un profil ou choisissez **Nouveau…** pour créer un profil.

    > [!NOTE]
    > Un profil de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.

    Si vous choisissez **Nouveau...**, un Assistant apparaît avec la page **Créer un profil de publication**. Vous pouvez utiliser cet Assistant pour importer un profil de publication à partir d’un site web d’hébergement comme Microsoft Azure ou créer un profil et ajouter votre serveur, vos informations d’identification et d’autres paramètres, comme décrit dans la procédure suivante.

    Pour plus d’informations sur l’importation et la création de profils de publication, reportez-vous à la rubrique [Création d’un profil de publication](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).

3. Sur la page **Publier votre complément**, cliquez sur le lien **Déployer votre projet web**.

    La boîte de dialogue **Publier** s’affiche. Pour plus d’informations sur l’utilisation de cet Assistant, reportez-vous à l’article relatif à la procédure de [déploiement d’un projet web à l’aide de On-Click Publishing dans Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>Création d’un package de votre complément avec Visual Studio 2017

Procédez comme suit pour créer un package de votre projet de complément à l’aide de Visual Studio 2017.

1. Sur la page **Publier votre complément**, cliquez sur le bouton permettant d’**empaqueter le complément**.

    Un Assistant s’affiche avec la page permettant d’**empaqueter le complément**.

2. Dans la liste déroulante **Où votre site web est-il hébergé ?**, sélectionnez ou saisissez l’URL du site web qui hébergera les fichiers de contenu de votre complément, puis cliquez sur **Terminer**.

    > [!IMPORTANT]
    > Les sites web Azure [!include[HTTPS guidance](../includes/https-guidance.md)] fournissent automatiquement un point de terminaison HTTPS.

    Visual Studio génère les fichiers nécessaires à la publication de votre complément, puis ouvre le dossier de sortie de publication.

Si vous prévoyez de soumettre votre complément à AppSource, vous pouvez cliquer sur le bouton **Effectuer la vérification de la validation** pour identifier les problèmes susceptibles d’empêcher votre complément d’être accepté. Vous devez corriger tous les problèmes avant d’envoyer votre complément au Store.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a>Voir aussi

- [Publier votre complément Office](../publish/publish.md)
- [Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-the-office-store)
