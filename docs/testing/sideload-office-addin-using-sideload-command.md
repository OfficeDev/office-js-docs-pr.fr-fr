---
title: Charger une version test des compléments Office à l'aide de la commande de chargement indépendant
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: c3b53a70b5696e422653350de18d99be16d1d597
ms.sourcegitcommit: 0d4d78e275249f0d4b6a6cf807b42b79890c3023
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21773593"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Chargez une version test des compléments Office à l'aide de la **commande de chargement indépendant**
 >[!NOTE]
>La méthode « npm run sideload » fonctionne uniquement pour les compléments Excel, Word et PowerPoint qui s’exécutent sur Windows ; et uniquement pour les projets de complément créés avec l’outil [**Yo Office**](https://github.com/OfficeDev/generator-office) et disposant d’un `sideload` script dans la section `scripts` du fichier package.json. (Les projets créés avec des versions antérieures de **Yo Office** ne disposent pas de ce script.) Si votre projet a été créé avec Visual Studio ou ne dispose pas du script sideload, vous pouvez en charger la version test sur Windows en suivant la méthode décrite dans [Chargement d’une version test de complément Office à partir d’un partage réseau](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour charger la version test de votre complément :
> 
> - [Chargement de version test des compléments Office dans Office Online](sideload-office-add-ins-for-testing.md)
> - [Chargement de version test des compléments Office sur iPad et Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Chargement de version test de compléments Outlook](../../../../outlook/add-insSideload Outlook Add-ins for testing)

1. Ouvrez une invite de commandes en tant qu’administrateur.

2. Modifiez les répertoires à la racine du dossier de projet du complément.

3. Exécutez la commande suivante pour démarrer une instance de serveur Web local sur le port 3000 afin de servir votre projet de complément :**« npm run start »**

4. Ouvrez une nouvelle invite de commandes en tant qu’administrateur.

5. Changez les répertoires à la racine du dossier de projet du complément.

6. Exécutez la commande suivante pour démarrer l'application hôte (par exemple Excel, Word) et enregistrez votre complément dans l'application hôte :**« npm run sideload »**

## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)