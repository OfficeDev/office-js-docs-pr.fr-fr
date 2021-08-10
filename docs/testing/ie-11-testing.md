---
title: Test d’Internet Explorer 11
description: Testez votre Office sur Internet Explorer 11.
ms.date: 06/18/2021
localization_priority: Normal
ms.openlocfilehash: 8579a37f1ea48d511010b8c55cfe9fad5aa6b41acee85b1da426e25083287655
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090127"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Tester votre Office sur Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer toujours utilisé dans les Office de recherche**
>
> Microsoft termine la prise en charge d’Internet Explorer, mais cela n’a pas d’incidence significative sur Office des modules. Certaines combinaisons de plateformes et de versions Office, y compris toutes les versions à achat unique jusqu’à Office 2019, continueront d’utiliser le contrôle webview qui est livré avec Internet Explorer 11 pour héberger des applications, comme expliqué dans les [navigateurs](../concepts/browsers-used-by-office-web-add-ins.md)utilisés par les applications Office . En outre, la prise en charge de ces combinaisons, et donc d’Internet Explorer, est toujours requise pour les applications soumises à [AppSource.](/office/dev/store/submit-to-appsource-via-partner-center) Deux choses *changent* :
>
> - AppSource ne teste plus les Office sur le Web l’aide d’Internet Explorer en tant que navigateur. Toutefois, AppSource teste toujours les combinaisons de plateforme et de Office *de bureau* qui utilisent Internet Explorer.
> - [L Script Lab’outil cessera](../overview/explore-with-script-lab.md) de fonctionner dans Internet Explorer en 2021.

Si vous envisagez de commercialiser votre application via AppSource ou si vous prévoyez de prendre en charge des versions antérieures de Windows et Office, votre application doit fonctionner dans le contrôle de navigateur in incorporer basé sur Internet Explorer 11 (IE11). Vous pouvez utiliser une ligne de commande pour passer de runtimes plus modernes utilisés par les modules de mise à l’essai à Internet Explorer 11 pour ce test. Pour plus d’informations sur les versions de Windows et Office utiliser le contrôle d’affichage web Internet Explorer 11, voir Navigateurs utilisés par les Office des [applications.](../concepts/browsers-used-by-office-web-add-ins.md)

> [!IMPORTANT]
> Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5. Si vous souhaitez utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou une ultérieure, vous disposez de deux options :
>
> - Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou version ultérieure JavaScript, ou dans TypeScript, puis compilez votre code en JavaScript ES5 à l’aide d’un compilateur tel que [celui-ci ou](https://babeljs.io/) [tsc.](https://www.typescriptlang.org/index.html)
> - Écrivez en JavaScript ECMAScript 2015 ou version ultérieure, mais chargez également une [bibliothèque polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) telle que [core-js](https://github.com/zloirock/core-js) qui permet à IE d’exécuter votre code.
>
> Pour plus d’informations sur ces options, voir [Support Internet Explorer 11](../develop/support-ie-11.md).
>
> Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement.

> [!NOTE]
> Pour tester votre add-in sur le navigateur Internet Explorer 11, ouvrez Office sur le Web dans Internet Explorer et chargez une version test [du module.](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))

Ces instructions supposent que vous avez déjà installé un projet Office Yo. Si vous ne l’avez pas encore fait, envisagez de lire un démarrage rapide, tel que [celui-ci pour Excel de recherche.](../quickstarts/excel-quickstart-jquery.md)

## <a name="switching-to-the-internet-explorer-11-webview"></a>Basculement vers le webview Internet Explorer 11

1. Créez un projet de générateur Office Yo. Peu importe le type de projet que vous sélectionnez, cet outil fonctionne avec tous les types de projets.

    > [!NOTE]
    > Si vous avez un projet existant et que vous souhaitez ajouter cet outil sans créer de nouveau projet, ignorez cette étape et passez à l’étape suivante. 

1. Dans le dossier racine de votre projet, exécutez la commande suivante dans la ligne de commande. Cet exemple suppose que le fichier manifeste de votre projet se trouve à la racine. Si ce n’est pas le cas, spécifiez le chemin d’accès relatif au fichier manifeste. Un message doit s’afficher dans la ligne de commande pour vous dire que le type d’affichage web est désormais définie sur IE.

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> Il n’est pas nécessaire d’utiliser cette commande, mais elle doit aider à déboguer la plupart des problèmes liés au runtime d’Internet Explorer 11. Pour une robustesse totale, vous devez tester l’utilisation d’ordinateurs avec différentes combinaisons de Windows 7, 8.1 et 10, ainsi que différentes versions de Office. Pour plus d’informations, voir [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).

### <a name="command-options"></a>Options de commande

La `office-addin-dev-settings webview` commande peut également prendre un certain nombre d’runtimes comme arguments :

- ie
- edge
- Valeur par défaut.

## <a name="see-also"></a>Voir aussi

* [Test et débogage de compléments Office](test-debug-office-add-ins.md)
* [Chargement de la version test des compléments Office](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Débogage des compléments avec les outils de développement sur Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
