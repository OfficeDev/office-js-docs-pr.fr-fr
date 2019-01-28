---
title: Chargement de versions test de compléments Office à l’aide de la commande sideload
description: ''
ms.date: 07/24/2018
localization_priority: Priority
ms.openlocfilehash: 2231e05d798dce4f4b5428627a3653ddcdecfc65
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387673"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Chargement de versions test de compléments Office pour les tester à l’aide de la **commande sideload**
 >[!NOTE]
>La méthode « npm run sideload » fonctionne uniquement pour les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows ; et uniquement pour les projets de complément qui ont été créés dans l’outil [**yo office** ](https://github.com/OfficeDev/generator-office)et qui ont un script `sideload` dans la section `scripts` du fichier package.json. (Les projets qui ont été créés dans les versions antérieures de **yo office** n’ont pas ce script non plus.) Si votre projet a été créé avec Visual Studio ou n’a pas le script sideload , vous pouvez charger une version test sur Windows en suivant la méthode décrite dans l’article relatif au [chargement de version test d’un complément Office à partir d’un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour charger la version test de votre complément :
> 
> - [Chargement de version test des compléments Office dans Office Online](sideload-office-add-ins-for-testing.md)
> - [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Chargement de version test des compléments Outlook pour les tester](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. Ouvrez une invite de commandes en tant qu’administrateur.

2. Modifiez les répertoires vers la racine du dossier de votre projet de complément.

3. Exécutez la commande suivante pour démarrer une instance du serveur web local sur le port 3000 et mettre en service votre projet de complément : « **npm exécuter début** »

4. Ouvrez une deuxième invite de commandes en tant qu’administrateur.

5. Modifiez les répertoires vers la racine du dossier de votre projet de complément.

6. Exécutez la commande suivante pour démarrer l’application hôte (par exemple, Excel, Word) et inscrire votre complément dans l’application hôte : « **npm run sideloadr** »

## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)
