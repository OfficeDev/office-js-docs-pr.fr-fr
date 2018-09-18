---
title: Charger une version test des compléments Office à l'aide de la commande de chargement indépendant
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 1ab0277493f2899adb479c2f24b1635a881af3cc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944040"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Chargez une version test des compléments Office à l'aide de la **commande de chargement indépendant**
 >[!NOTE]
>La méthode « npm exécuter sideload » ne fonctionne que pour les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows ; et uniquement pour les projets de compléments qui ont été créés avec l'outil [**yo office**](https://github.com/OfficeDev/generator-office) et qui ont un script `sideload` dans la section `scripts` du fichier package.json. (Les projets qui ont été créées avec les versions antérieures de **yo office** n’ont pas non plus ce script.) Si votre projet a été créé avec Visual Studio ou n’a pas le script sideload, vous pouvez le charger en version test sur Windows avec la méthode décrite dans [Chargement de la version test d'un complément Office depuis un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Si ce n'est pas un complément Word, Excel ou PowerPoint sous Windows que vous testez, consultez une des rubriques suivantes pour charger la version test de votre complément :
> 
> - [Chargement de version test des compléments Office dans Office Online](sideload-office-add-ins-for-testing.md)
> - [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Chargement de version test des compléments Outlook](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. Ouvrez une invite de commandes en tant qu’administrateur.

2. Modifiez les répertoires à la racine du dossier de projet du complément.

3. Exécutez la commande suivante pour démarrer une instance de serveur Web local sur le port 3000 afin de servir votre projet de complément :**« npm run start »**

4. Ouvrez une nouvelle invite de commandes en tant qu’administrateur.

5. Changez les répertoires à la racine du dossier de projet du complément.

6. Exécutez la commande suivante pour démarrer l'application hôte (par exemple Excel, Word) et enregistrez votre complément dans l'application hôte :**« npm run sideload »**

## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)