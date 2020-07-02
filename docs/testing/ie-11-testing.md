---
ms.date: 05/16/2020
description: Testez votre complément Office à l’aide d’Internet Explorer 11.
title: Test Internet Explorer 11
localization_priority: Normal
ms.openlocfilehash: 1d6852d08308088a020e86ce7f5ab9cfdb9ab978
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006436"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a>Tester votre complément Office à l’aide d’Internet Explorer 11

En fonction des spécifications de votre complément, vous pouvez envisager de prendre en charge des versions antérieures de Windows et d’Office, qui nécessitent des tests sur Internet Explorer 11. Cela est souvent nécessaire dans le cadre de l’envoi de votre complément à AppSource. Vous pouvez utiliser les outils de ligne de commande suivants pour basculer d’autres runtimes modernes utilisés par les compléments vers le runtime Internet Explorer 11 pour ce test.

## <a name="pre-requisites"></a>Conditions préalables

- [Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))
- Éditeur de code. Nous recommandons [Visual Studio code](https://code.visualstudio.com/)
- [Faire partie du programme Office Insider](https://insider.office.com)

Ces instructions supposent que vous avez configuré un projet de générateur Office Yo avant. Si vous n’avez pas encore fait cela, envisagez de lire un démarrage rapide, tel que celui- [ci pour les compléments Excel](../quickstarts/excel-quickstart-jquery.md).

## <a name="using-ie11-tooling"></a>Utilisation des outils IE11

1. Créez un projet de générateur Office Yo. Quel que soit le type de projet que vous sélectionnez, ces outils fonctionnent avec tous les types de projets.

> ! Note Si vous disposez d’un projet existant et que vous souhaitez ajouter cet outil sans créer de nouveau projet, ignorez cette étape et passez à l’étape suivante. 

2. Dans le dossier racine de votre nouveau projet, exécutez la commande suivante dans la ligne de commande :

```command&nbsp;line
npx office-addin-dev-settings webview manifest.xml ie
```
Vous devriez voir une remarque dans la ligne de commande que le type d’affichage Web est maintenant défini sur Internet Explorer.

> ! TETE Il n’est pas nécessaire d’utiliser cet outil, mais cela devrait vous aider à déboguer la majorité des problèmes liés à Internet Explorer 11 Runtime. Pour une robustesse totale, vous devez tester à l’aide d’un ordinateur sur lequel une copie de Windows 7 et Office 2013 est installée.

## <a name="command-settings"></a>Paramètres de la commande

Si vous avez un chemin d’accès de manifeste différent, spécifiez-le dans la commande, comme indiqué dans l’exemple suivant :

`npx office-addin-dev-settings webview [path to your manifest] ie`

La `office-addin-dev-settings webview` commande peut également prendre un certain nombre d’exécutions en tant qu’arguments :

- échange
- cadre
- Valeur par défaut.

## <a name="see-also"></a>Voir aussi
* [Test et débogage de compléments Office](test-debug-office-add-ins.md)
* [Chargement de la version test des compléments Office](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Débogage des compléments avec les outils de développement sur Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
