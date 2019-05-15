---
title: Chargement de versions test de compléments Office à l’aide de la commande sideload
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: 69d39c2736312653b5a362aefccd41629e6e3555
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619076"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Chargement indépendant de compléments Office pour les tester à l’aide de la commande sideload
 
> [!NOTE]
> La technique de chargement indépendant décrite dans cet article est uniquement valide pour :
> 
> - Les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows.
> 
> - Les projets de complément créés avec le [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office) et disposant d’un script `sideload` dans la section `scripts` du fichier package.json. (Ce script n’est pas présent dans les projets créés avec d’anciennes versions du générateur Yeoman pour compléments Office).
 
Pour charger indépendamment votre complément à l’aide du script `sideload` fourni par le générateur Yeoman pour compléments Office, procédez comme suit :

1. Ouvrez une invite de commandes en tant qu’administrateur.

2. Modifiez les répertoires vers la racine du dossier de votre projet de complément.

3. Exécutez la commande suivante pour démarrer une instance du serveur web local sur le port 3000 et mettre en service votre projet de complément : `npm run start`

4. Ouvrez une deuxième invite de commandes en tant qu’administrateur.

5. Modifiez les répertoires vers la racine du dossier de votre projet de complément.

6. Exécutez la commande suivante pour démarrer l’application hôte (par exemple, Excel, Word) et inscrire votre complément dans l’application hôte : `npm run sideload`

Si votre projet de complément a été créé avec Visual Studio ou n’a pas le script sideload , vous pouvez le charger indépendamment sur Windows en suivant la méthode décrite dans l’article relatif au [chargement indépendant d’un complément Office à partir d’un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour plus d’informations sur le chargement indépendant de votre complément :
 
- [Chargement de version test des compléments Office dans Office Online pour les tester](sideload-office-add-ins-for-testing.md)
- [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook pour les tester](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)
