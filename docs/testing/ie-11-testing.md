---
title: Test d’Internet Explorer 11
description: Testez votre Office sur Internet Explorer 11.
ms.date: 10/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: cfa6a35565fdca28eab9734ccde9fc8fbb2e8270
ms.sourcegitcommit: a37be80cf47a37c85b7f5cab216c160f4e905474
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/09/2021
ms.locfileid: "60250517"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Tester votre Office sur Internet Explorer 11

> [!IMPORTANT]
> **Internet Explorer toujours utilisé dans les Office de recherche**
>
> Microsoft termine la prise en charge d’Internet Explorer, mais cela n’a pas d’incidence significative sur Office des modules. Certaines combinaisons de plateformes et de versions Office, y compris les versions d’achat unique jusqu’à Office 2019, continueront d’utiliser le contrôle webview qui est livré avec Internet Explorer 11 pour héberger des applications, comme expliqué dans les navigateurs utilisés par les applications [Office.](../concepts/browsers-used-by-office-web-add-ins.md) En outre, la prise en charge de ces combinaisons, et donc d’Internet Explorer, est toujours requise pour les applications soumises à [AppSource.](/office/dev/store/submit-to-appsource-via-partner-center) Deux choses *changent* :
>
> - Office sur le Web ne s’ouvre plus dans Internet Explorer. Par conséquent, AppSource ne teste plus les Office sur le Web à l’aide d’Internet Explorer en tant que navigateur. Toutefois, AppSource teste toujours les combinaisons de plateforme et de Office *de bureau* qui utilisent Internet Explorer.
> - [L Script Lab ne prend](../overview/explore-with-script-lab.md) plus en charge Internet Explorer.

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
> Office sur le Web ne peut pas être ouvert dans Internet Explorer 11, vous ne pouvez pas (et n’avez pas besoin de) tester votre module sur Office sur le Web avec Internet Explorer.

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
> Il n’est pas nécessaire d’utiliser cette commande, mais elle doit aider à déboguer la plupart des problèmes liés au runtime d’Internet Explorer 11. Pour une robustesse totale, vous devez tester l’utilisation d’ordinateurs avec différentes combinaisons de Windows 7, 8.1, 10 et 11, ainsi que différentes versions de Office. Pour plus d’informations, voir [Navigateurs](../concepts/browsers-used-by-office-web-add-ins.md) utilisés par les Office et comment revenir à une version antérieure de [Office](https://support.microsoft.com/topic/2bd5c457-a917-d57e-35a1-f709e3dda841).

### <a name="command-options"></a>Options de commande

La `office-addin-dev-settings webview` commande peut également prendre un certain nombre d’runtimes comme arguments :

- ie
- edge
- Valeur par défaut.

## <a name="see-also"></a>Voir aussi

* [Test et débogage de compléments Office](test-debug-office-add-ins.md)
* [Chargement de la version test des compléments Office](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Déboguer des applications à l’aide des outils de développement sur Windows](debug-add-ins-using-f12-developer-tools-on-windows.md)
* [Attacher un débogueur à partir du volet Office](attach-debugger-from-task-pane.md)
