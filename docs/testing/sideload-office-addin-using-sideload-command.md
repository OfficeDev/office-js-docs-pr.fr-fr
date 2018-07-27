---
title: Charger une version test des compléments Office à l'aide de la commande de chargement indépendant
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: e831a1dfbc31ecf06c8b2d78dc1e9a8a4c9dcf01
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/27/2018
ms.locfileid: "21279359"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Chargez une version test des compléments Office à l'aide de la **commande de chargement indépendant**
 >[!NOTE]
>La méthode « npm run sideload » ne fonctionne que pour les compléments Excel, Word et PowerPoint.

1. Ouvrez une invite de commandes en tant qu’administrateur.

2. Modifiez les répertoires à la racine du dossier de projet du complément.

3. Exécutez la commande suivante pour démarrer une instance de serveur Web local sur le port 3000 afin de servir votre projet de complément :**« npm run start »**

4. Ouvrez une nouvelle invite de commandes en tant qu’administrateur.

5. Changez les répertoires à la racine du dossier de projet du complément.

6. Exécutez la commande suivante pour démarrer l'application hôte (par exemple Excel, Word) et enregistrez votre complément dans l'application hôte :**« npm run sideload »**

## <a name="see-also"></a>Voir aussi

- [Valider et résoudre des problèmes avec votre manifeste](troubleshoot-manifest.md)
- [Publier votre complément Office](../publish/publish.md)