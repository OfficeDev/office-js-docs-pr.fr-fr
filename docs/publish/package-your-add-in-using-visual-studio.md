---
title: Publier votre complément à l’aide de Visual Studio
description: Déploiement de votre projet web et création d’un package de votre complément à l’aide de Visual Studio 2019.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 78b80e0c6a595f83f3a8cdde1db806a7612036ea
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950718"
---
# <a name="publish-your-add-in-using-visual-studio"></a>Publier votre complément à l’aide de Visual Studio

Votre package de complément Office contient un [fichier manifeste](../develop/add-in-manifests.md) XML que vous allez utiliser pour publier le complément. Vous devez publier les fichiers d’application web de votre projet séparément. Cet article décrit le déploiement de votre projet web et création d’un package de votre complément à l’aide de Visual Studio 2019.

> [!NOTE]
> Pour plus d’informations sur la publication d’un complément Office que vous avez créé à l’aide du générateur Yeoman et développé avec Visual Studio Code ou un autre éditeur, voir [Publier un complément développé avec Visual Studio Code](publish-add-in-vs-code.md).

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a>Pour déployer votre projet web à l’aide de Visual Studio 2019

Réalisez les étapes suivantes pour déployer votre projet Web à l'aide de Visual Studio 2019.

1. Depuis l’onglet **Build**, sélectionnez **Publier [nom de votre complément]**.

2. Dans la fenêtre **Choisir une cible de publication **, sélectionnez une des options pour publier sur votre cible préférée. Chaque cible de publication nécessite que vous incluiez plus d'informations pour commencer, comme l'emplacement d'une machine virtuelle Azure ou d'un emplacement de dossier. Une fois que vous avez spécifié un emplacement de publication et renseigné toutes les informations requises, sélectionnez **Publier**

    > [!NOTE]
    > Le choix d’une cible de publication indique le serveur sur lequel vous effectuez le déploiement, les informations d’identification nécessaires pour se connecter au serveur, les bases de données à déployer, ainsi que d’autres options de déploiement.

3. Pour plus d’informations sur les étapes de déploiement de chaque option cible de publication, voir [Premier aperçu du déploiement dans Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a>Pour créer un package et publier votre complément à l’aide d’IIS, de FTP ou du déploiement Web à l’aide de Visual Studio 2019

Procédez comme suit pour créer un package de votre complément à l’aide de Visual Studio 2019.

1. Depuis l’onglet **Build**, sélectionnez **Publier [nom de votre complément]**.
2. Dans la fenêtre **Choisir une cible de publication**, choisissez **IIS, FTP, etc.** et sélectionnez **Configurer**. Sélectionnez ensuite **Publier**.
3. Un assistant s’affiche pour vous guider tout au long du processus. Assurez-vous que la méthode de publication est votre méthode préférée, telle que Web Deploy.
4. Dans la zone **URL de destination**, entrez l'URL du site Web qui hébergera les fichiers de contenu de votre complément, puis sélectionnez **Suivant**. Si vous prévoyez de soumettre votre complément à AppSource, vous pouvez choisir le bouton **Valider la connexion** pour identifier tout problème susceptible d'empêcher votre complément d'être accepté. Vous devez corriger tous les problèmes avant d’envoyer votre complément au Store.
5. Confirmez tous les paramètres souhaités, y compris les **Options de publication de fichiers**, puis sélectionnez **Enregistrer**.

    > [!IMPORTANT]
    > Les sites web Azure [!include[HTTPS guidance](../includes/https-guidance.md)] fournissent automatiquement un point de terminaison HTTPS.

Vous pouvez désormais télécharger votre manifeste XML à l’emplacement approprié pour [publier votre complément](../publish/publish.md). Le manifeste XML se trouve dans `OfficeAppManifests` dans le dossier `app.publish`. Par exemple :

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a>Voir aussi

- [Publier votre complément Office](../publish/publish.md)
- [Mise à disposition de vos solutions sur AppSource et dans Office](/office/dev/store/submit-to-the-office-store)
